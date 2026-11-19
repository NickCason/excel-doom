"""Minimal WebAssembly binary parser plus a validation-style walk that annotates
every instruction with stack height, operand type and resolved branch targets."""
import struct

I32, I64, F32, F64 = 0x7F, 0x7E, 0x7D, 0x7C
TYNAME = {I32: 'i32', I64: 'i64', F32: 'f32', F64: 'f64'}

def leb_u(b, i):
    r = 0; s = 0
    while True:
        x = b[i]; i += 1
        r |= (x & 0x7F) << s; s += 7
        if not x & 0x80:
            return r, i

def leb_s(b, i):
    r = 0; s = 0
    while True:
        x = b[i]; i += 1
        r |= (x & 0x7F) << s; s += 7
        if not x & 0x80:
            if x & 0x40:
                r -= 1 << s
            return r, i

BLOCK, LOOP, IF, ELSE, END = 0x02, 0x03, 0x04, 0x05, 0x0B

def decode_body(b, i, end):
    out = []
    while i < end:
        op = b[i]; i += 1
        if op in (BLOCK, LOOP, IF):
            if b[i] == 0x40:
                bt = (); i += 1
            elif b[i] in (I32, I64, F32, F64):
                bt = (b[i],); i += 1
            else:
                raise NotImplementedError('multi-value block type')
            out.append((op, bt))
        elif op in (0x0C, 0x0D, 0x10, 0x20, 0x21, 0x22, 0x23, 0x24):
            v, i = leb_u(b, i); out.append((op, v))
        elif op == 0x0E:
            n, i = leb_u(b, i); t = []
            for _ in range(n):
                v, i = leb_u(b, i); t.append(v)
            d, i = leb_u(b, i); out.append((op, t, d))
        elif op == 0x11:
            ti, i = leb_u(b, i); tb, i = leb_u(b, i); out.append((op, ti, tb))
        elif 0x28 <= op <= 0x3E:
            al, i = leb_u(b, i); off, i = leb_u(b, i); out.append((op, al, off))
        elif op in (0x3F, 0x40):
            i += 1; out.append((op,))
        elif op == 0x41 or op == 0x42:
            v, i = leb_s(b, i); out.append((op, v))
        elif op == 0x43:
            out.append((op, struct.unpack_from('<f', b, i)[0])); i += 4
        elif op == 0x44:
            out.append((op, struct.unpack_from('<d', b, i)[0])); i += 8
        elif op == 0xFC:
            sub, i = leb_u(b, i)
            if sub == 8:
                d, i = leb_u(b, i); i += 1; out.append((0xFC00 | sub, d))
            elif sub == 9:
                d, i = leb_u(b, i); out.append((0xFC00 | sub, d))
            elif sub == 10:
                i += 2; out.append((0xFC00 | sub,))
            elif sub == 11:
                i += 1; out.append((0xFC00 | sub,))
            else:
                out.append((0xFC00 | sub,))
        else:
            out.append((op,))
    return out

class Module:
    pass

def const_expr(b, i):
    op = b[i]; i += 1
    if op == 0x41 or op == 0x42:
        v, i = leb_s(b, i)
    elif op == 0x23:
        v, i = leb_u(b, i); v = ('global', v)
    else:
        raise NotImplementedError('const expr op %#x' % op)
    assert b[i] == 0x0B
    return v, i + 1

def parse(path):
    b = open(path, 'rb').read()
    assert b[:8] == b'\0asm\x01\0\0\0'
    m = Module()
    m.types = []; m.imports = []; m.func_types = []; m.globals = []; m.exports = {}
    m.elems = []; m.codes = []; m.datas = []; m.mem = (0, None); m.names = {}; m.table_size = 0
    i = 8
    while i < len(b):
        sid = b[i]; i += 1
        size, i = leb_u(b, i)
        end = i + size
        p = i
        if sid == 1:
            n, p = leb_u(b, p)
            for _ in range(n):
                assert b[p] == 0x60; p += 1
                np_, p = leb_u(b, p); params = tuple(b[p:p + np_]); p += np_
                nr, p = leb_u(b, p); results = tuple(b[p:p + nr]); p += nr
                m.types.append((params, results))
        elif sid == 2:
            n, p = leb_u(b, p)
            for _ in range(n):
                l, p = leb_u(b, p); mod = b[p:p + l].decode(); p += l
                l, p = leb_u(b, p); name = b[p:p + l].decode(); p += l
                kind = b[p]; p += 1
                if kind == 0:
                    ti, p = leb_u(b, p); m.imports.append((mod, name, ti))
                else:
                    raise NotImplementedError('non-function import')
        elif sid == 3:
            n, p = leb_u(b, p)
            for _ in range(n):
                ti, p = leb_u(b, p); m.func_types.append(ti)
        elif sid == 4:
            n, p = leb_u(b, p)
            for _ in range(n):
                p += 1
                fl, p = leb_u(b, p); mn, p = leb_u(b, p)
                if fl & 1:
                    _, p = leb_u(b, p)
                m.table_size = mn
        elif sid == 5:
            n, p = leb_u(b, p)
            fl, p = leb_u(b, p); mn, p = leb_u(b, p); mx = None
            if fl & 1:
                mx, p = leb_u(b, p)
            m.mem = (mn, mx)
        elif sid == 6:
            n, p = leb_u(b, p)
            for _ in range(n):
                t = b[p]; mut = b[p + 1]; p += 2
                v, p = const_expr(b, p)
                m.globals.append((t, mut, v))
        elif sid == 7:
            n, p = leb_u(b, p)
            for _ in range(n):
                l, p = leb_u(b, p); name = b[p:p + l].decode(); p += l
                kind = b[p]; p += 1
                idx, p = leb_u(b, p)
                m.exports[name] = (kind, idx)
        elif sid == 9:
            n, p = leb_u(b, p)
            for _ in range(n):
                fl, p = leb_u(b, p)
                if fl != 0:
                    raise NotImplementedError('element segment flag %d' % fl)
                off, p = const_expr(b, p)
                c, p = leb_u(b, p); fs = []
                for _ in range(c):
                    f, p = leb_u(b, p); fs.append(f)
                m.elems.append((off, fs))
        elif sid == 10:
            n, p = leb_u(b, p)
            for _ in range(n):
                sz, p = leb_u(b, p); fend = p + sz
                nl, p = leb_u(b, p); locs = []
                for _ in range(nl):
                    c, p = leb_u(b, p); locs += [b[p]] * c; p += 1
                m.codes.append((locs, decode_body(b, p, fend)))
                p = fend
        elif sid == 11:
            n, p = leb_u(b, p)
            for _ in range(n):
                fl, p = leb_u(b, p)
                if fl == 0:
                    off, p = const_expr(b, p)
                elif fl == 1:
                    off = None
                else:
                    raise NotImplementedError('data segment flag %d' % fl)
                l, p = leb_u(b, p)
                m.datas.append((off, b[p:p + l])); p += l
        elif sid == 0:
            l, p = leb_u(b, p); nm = b[p:p + l]; p += l
            if nm == b'name':
                while p < end:
                    sub = b[p]; p += 1
                    ss, p = leb_u(b, p); se = p + ss
                    if sub == 1:
                        c, q = leb_u(b, p)
                        for _ in range(c):
                            idx, q = leb_u(b, q); l, q = leb_u(b, q)
                            m.names[idx] = b[q:q + l].decode(errors='replace'); q += l
                    p = se
        i = end
    m.nimports = len(m.imports)
    return m

def func_type(m, fidx):
    if fidx < m.nimports:
        return m.types[m.imports[fidx][2]]
    return m.types[m.func_types[fidx - m.nimports]]

# stack signatures for plain numeric ops: op -> (pops, pushes)
SIG = {}
def _sig(ops, pops, push):
    for o in ops:
        SIG[o] = (pops, push)
_sig([0x45], (I32,), (I32,)); _sig(range(0x46, 0x50), (I32, I32), (I32,))
_sig([0x50], (I64,), (I32,)); _sig(range(0x51, 0x5B), (I64, I64), (I32,))
_sig(range(0x5B, 0x61), (F32, F32), (I32,)); _sig(range(0x61, 0x67), (F64, F64), (I32,))
_sig(range(0x67, 0x6A), (I32,), (I32,)); _sig(range(0x6A, 0x79), (I32, I32), (I32,))
_sig(range(0x79, 0x7C), (I64,), (I64,)); _sig(range(0x7C, 0x8B), (I64, I64), (I64,))
_sig(range(0x8B, 0x92), (F32,), (F32,)); _sig(range(0x92, 0x99), (F32, F32), (F32,))
_sig(range(0x99, 0xA0), (F64,), (F64,)); _sig(range(0xA0, 0xA7), (F64, F64), (F64,))
_sig([0xA7], (I64,), (I32,)); _sig([0xA8, 0xA9], (F32,), (I32,)); _sig([0xAA, 0xAB], (F64,), (I32,))
_sig([0xAC, 0xAD], (I32,), (I64,)); _sig([0xAE, 0xAF], (F32,), (I64,)); _sig([0xB0, 0xB1], (F64,), (I64,))
_sig([0xB2, 0xB3], (I32,), (F32,)); _sig([0xB4, 0xB5], (I64,), (F32,)); _sig([0xB6], (F64,), (F32,))
_sig([0xB7, 0xB8], (I32,), (F64,)); _sig([0xB9, 0xBA], (I64,), (F64,)); _sig([0xBB], (F32,), (F64,))
_sig([0xBC], (F32,), (I32,)); _sig([0xBD], (F64,), (I64,)); _sig([0xBE], (I32,), (F32,)); _sig([0xBF], (I64,), (F64,))
_sig([0xC0, 0xC1], (I32,), (I32,)); _sig([0xC2, 0xC3, 0xC4], (I64,), (I64,))
for _o in range(0xFC00, 0xFC08):
    SIG[_o] = (((F32,), (F32,), (F64,), (F64,), (F32,), (F32,), (F64,), (F64,))[_o & 7], ((I32,) if (_o & 7) < 4 else (I64,)))
SIG[0xFC0A] = ((I32, I32, I32), ()); SIG[0xFC0B] = ((I32, I32, I32), ())
LOAD_T = {0x28: I32, 0x29: I64, 0x2A: F32, 0x2B: F64, 0x2C: I32, 0x2D: I32, 0x2E: I32, 0x2F: I32,
          0x30: I64, 0x31: I64, 0x32: I64, 0x33: I64, 0x34: I64, 0x35: I64}
STORE_T = {0x36: I32, 0x37: I64, 0x38: F32, 0x39: F64, 0x3A: I32, 0x3B: I32, 0x3C: I64, 0x3D: I64, 0x3E: I64}

def analyze(m, k):
    """Annotate defined function k. Returns (local_types, result_types, ann) where ann[i] is a
    dict: h (height before), dead, and per-op extras (ty, br targets, frame matches)."""
    params, results = m.types[m.func_types[k]]
    locs, ins = m.codes[k]
    ltypes = list(params) + list(locs)
    n = len(ins)
    ann = [dict() for _ in range(n)]
    # match structure
    st = []
    for i, x in enumerate(ins):
        if x[0] in (BLOCK, LOOP, IF):
            st.append(i); ann[i]['else'] = None
        elif x[0] == ELSE:
            ann[st[-1]]['else'] = i; ann[i]['start'] = st[-1]
        elif x[0] == END:
            if st:
                s = st.pop(); ann[s]['end'] = i; ann[i]['start'] = s
            else:
                ann[i]['start'] = None
    stack = []
    func_frame = dict(kind='func', start=None, height=0, results=results, live=True)
    ctrl = [func_frame]
    dead = False
    def target(depth):
        f = ctrl[-1 - depth]
        if f['kind'] == 'loop':
            return dict(kind='loop', start=f['start'], height=f['height'], arity=())
        return dict(kind=f['kind'], start=f['start'], height=f['height'], arity=f['results'])
    for i, x in enumerate(ins):
        op = x[0]; a = ann[i]
        a['h'] = len(stack); a['dead'] = dead
        if dead:
            if op in (BLOCK, LOOP, IF):
                ctrl.append(dict(kind='deadnest', start=i, height=len(stack), results=(), live=False))
            elif op == ELSE:
                f = ctrl[-1]
                if f['live']:
                    del stack[f['height']:]; dead = False; a['revive'] = True
            elif op == END:
                f = ctrl.pop()
                if f['live']:
                    del stack[f['height']:]; stack += list(f['results']); dead = False; a['revive'] = True
            continue
        if op in (BLOCK, LOOP, IF):
            if op == IF:
                stack.pop()
            ctrl.append(dict(kind={BLOCK: 'block', LOOP: 'loop', IF: 'if'}[op], start=i, height=len(stack), results=x[1], live=True))
        elif op == ELSE:
            f = ctrl[-1]; del stack[f['height']:]
        elif op == END:
            f = ctrl.pop(); del stack[f['height']:]; stack += list(f['results'])
        elif op == 0x0C:
            a['br'] = target(x[1]); dead = True
        elif op == 0x0D:
            stack.pop(); a['br'] = target(x[1])
        elif op == 0x0E:
            stack.pop(); a['brs'] = [target(d) for d in x[1]]; a['brd'] = target(x[2]); dead = True
        elif op == 0x0F or op == 0x00:
            dead = True
        elif op == 0x01:
            pass
        elif op == 0x10:
            p_, r_ = func_type(m, x[1]); del stack[len(stack) - len(p_):]; stack += list(r_)
        elif op == 0x11:
            stack.pop(); p_, r_ = m.types[x[1]]; del stack[len(stack) - len(p_):]; stack += list(r_)
        elif op == 0x1A:
            a['ty'] = stack.pop()
        elif op == 0x1B:
            stack.pop(); a['ty'] = stack.pop()
        elif op == 0x20:
            stack.append(ltypes[x[1]])
        elif op == 0x21:
            stack.pop()
        elif op == 0x22:
            pass
        elif op == 0x23:
            stack.append(m.globals[x[1]][0])
        elif op == 0x24:
            stack.pop()
        elif op in LOAD_T:
            stack.pop(); stack.append(LOAD_T[op])
        elif op in STORE_T:
            stack.pop(); stack.pop()
        elif op == 0x3F:
            stack.append(I32)
        elif op == 0x40:
            pass
        elif op == 0x41:
            stack.append(I32)
        elif op == 0x42:
            stack.append(I64)
        elif op == 0x43:
            stack.append(F32)
        elif op == 0x44:
            stack.append(F64)
        elif op in SIG:
            pops, push = SIG[op]
            del stack[len(stack) - len(pops):]; stack += list(push)
        else:
            raise NotImplementedError('op %#x' % op)
    return ltypes, results, ann

OPNAMES = {0x00: 'unreachable', 0x01: 'nop', 0x02: 'block', 0x03: 'loop', 0x04: 'if', 0x05: 'else', 0x0B: 'end',
           0x0C: 'br', 0x0D: 'br_if', 0x0E: 'br_table', 0x0F: 'return', 0x10: 'call', 0x11: 'call_indirect',
           0x1A: 'drop', 0x1B: 'select', 0x20: 'local.get', 0x21: 'local.set', 0x22: 'local.tee',
           0x23: 'global.get', 0x24: 'global.set', 0x3F: 'memory.size', 0x40: 'memory.grow',
           0x41: 'i32.const', 0x42: 'i64.const', 0x43: 'f32.const', 0x44: 'f64.const',
           0xFC0A: 'memory.copy', 0xFC0B: 'memory.fill'}

def opclass(op):
    if op in OPNAMES: return OPNAMES[op]
    if op in LOAD_T: return TYNAME[LOAD_T[op]] + '.load*'
    if op in STORE_T: return TYNAME[STORE_T[op]] + '.store*'
    if 0x45 <= op <= 0x4F or 0x67 <= op <= 0x78 or op in (0xC0, 0xC1): return 'i32 arith/cmp'
    if 0x50 <= op <= 0x5A or 0x79 <= op <= 0x8A or op in (0xC2, 0xC3, 0xC4): return 'i64 arith/cmp'
    if 0x5B <= op <= 0x66 or 0x8B <= op <= 0xA6: return 'float arith/cmp'
    if 0xA7 <= op <= 0xBF: return 'conversion' + (' (float)' if op not in (0xA7, 0xAC, 0xAD) else ' (int)')
    if 0xFC00 <= op <= 0xFC07: return 'float trunc_sat'
    return 'op %#x' % op
