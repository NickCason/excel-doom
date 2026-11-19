"""WebAssembly -> VBA translator, full-module version (integers, 64-bit, floats, bulk
memory, imports). Output sticks to a small VBA subset that vbaemu.py can also run."""
import re, struct
import wasm
from wasm import I32, I64, F32, F64, BLOCK, LOOP, IF, ELSE, END, func_type

VT = {I32: 'Long', I64: 'LongLong', F32: 'Single', F64: 'Double'}
SLOT = {I32: 's', I64: 'j', F32: 'e', F64: 'd'}
W32 = 'CLng(((%s And 4294967295^) Xor 2147483648^) - 2147483648^)'

def s32(v):
    v &= 0xFFFFFFFF
    return v - (1 << 32) if v >= 0x80000000 else v

def s64(v):
    v &= (1 << 64) - 1
    return v - (1 << 64) if v >= 1 << 63 else v

def lit32(v):
    v = s32(v)
    if v == -2147483648:
        return '&H80000000'
    return '%d&' % v if v >= 0 else '(%d&)' % v

def lit64(v):
    v = s64(v)
    if v == -(1 << 63):
        return 'P2(63)'
    return '%d^' % v if v >= 0 else '(%d^)' % v

class E:
    __slots__ = ('x', 't', 'deps', 'mem', 'kind', 'c', 'bits')
    def __init__(self, x, t, deps=frozenset(), mem=False, kind='v', c=None, bits=None):
        self.x = x; self.t = t; self.deps = deps; self.mem = mem; self.kind = kind; self.c = c; self.bits = bits
    def simple(self):
        return self.c is not None or self.x.isidentifier()

CMP_S = {0x46: '=', 0x47: '<>', 0x48: '<', 0x4A: '>', 0x4C: '<=', 0x4E: '>='}
CMP_U = {0x49: '<', 0x4B: '>', 0x4D: '<=', 0x4F: '>='}
CMP64_S = {0x51: '=', 0x52: '<>', 0x53: '<', 0x55: '>', 0x57: '<=', 0x59: '>='}
CMP64_U = {0x54: '<', 0x56: '>', 0x58: '<=', 0x5A: '>='}
FCMP = {0: '=', 1: '<>', 2: '<', 3: '>', 4: '<=', 5: '>='}
FUN1 = {0: 'Abs(%s)', 1: '(-%s)', 2: 'FCeil(%s)', 3: 'FFloor(%s)', 4: 'Fix(%s)', 5: 'FNearest(%s)', 6: 'Sqr(%s)'}
FBIN = {7: '(%s + %s)', 8: '(%s - %s)', 9: '(%s * %s)', 10: '(%s / %s)', 11: 'FMin(%s, %s)', 12: 'FMax(%s, %s)', 13: 'FCopySign(%s, %s)'}

class Unsupported(Exception):
    pass

class Fn:
    def __init__(self, T, k):
        self.T = T; self.m = T.m; self.k = k
        self.ltypes, self.results, self.ann = wasm.analyze(self.m, k)
        self.ins = self.m.codes[k][1]
        self.nparams = len(self.m.types[self.m.func_types[k]][0])
        self.name = 'f%d' % (k + self.m.nimports)
        self.out = []; self.st = []; self.used = set(); self.frames = {}; self.temps = set()

    def emit(self, s):
        self.out.append(s)

    def canon(self, d, t):
        n = '%s%d' % (SLOT[t], d)
        self.used.add((n, t))
        return n

    def push(self, x, t, deps=frozenset(), mem=False, kind='v', c=None, bits=None):
        self.st.append(E(x, t, deps, mem, kind, c, bits))
        if len(x) > 160:
            self.flush()

    def val(self, e):
        return '((%s) And 1&)' % e.x if e.kind == 'b' else e.x

    def cond(self, e):
        return e.x if e.kind == 'b' else '(%s <> 0)' % e.x

    def flush(self):
        for d, e in enumerate(self.st):
            cn = self.canon(d, e.t)
            if e.x != cn:
                self.emit('%s = %s' % (cn, self.val(e)))
                self.st[d] = E(cn, e.t)

    def hazard_local(self, name):
        if any(name in e.deps for e in self.st):
            self.flush()

    def hazard_mem(self):
        if any(e.mem or 'G' in e.deps for e in self.st):
            self.flush()

    def result(self, t):
        self.flush()
        n = self.canon(len(self.st), t)
        self.st.append(E(n, t))
        return n

    def reset(self, types):
        self.st = [E(self.canon(d, t), t) for d, t in enumerate(types)]

    def tmp(self, n, t='Long'):
        self.temps.add((n, t)); return n

    # ---------------------------------------------------------------- memory
    def addr(self, base, off):
        if base.c is not None:
            return None, (base.c + off) & 0xFFFFFFFF
        b = self.val(base)
        return (b if off == 0 else '(%s + %d&)' % (b, off)), None

    def widx(self, x, ca):
        """Word index expression for a trusted word-aligned address."""
        if x is None:
            if ca % 4 == 0:
                return '%d' % (ca // 4)
            x = lit32(ca)
        return ('AL4(%s)' if self.T.checked else '%s \\ 4') % x

    def load(self, op, align, off):
        base = self.st.pop()
        x, ca = self.addr(base, off)
        ax = x if x is not None else lit32(ca)
        deps = base.deps
        P = lambda e, t=I32, **kw: self.push(e, t, deps, True, **kw)
        if op == 0x28:
            if align >= 2 and not (x is None and ca % 4):
                return P('M(%s)' % self.widx(x, ca))
            return P('LD32(%s)' % ax)
        if op in (0x2D, 0x2C, 0x31, 0x30):
            self.tmp('a')
            body = '((M(a \\ 4) And MK(a And 3)) \\ DV(a And 3)) And 255'
            if op in (0x2C, 0x30):
                body = '((%s) Xor 128) - 128' % body
            t = I64 if op in (0x30, 0x31) else I32
            r = self.result(t)
            self.emit('a = %s: %s = %s' % (ax, r, body)); return
        if op in (0x2F, 0x2E, 0x33, 0x32):
            e = ('LD16A(%s)' if (align >= 1 and self.T.checked) else 'LD16(%s)') % ax
            if align >= 1 and not self.T.checked and x is not None:
                self.tmp('a')
                body = '((M(a \\ 4) And MH(a And 2)) \\ DV(a And 2)) And 65535&'
                if op in (0x2E, 0x32):
                    body = '((%s) Xor 32768&) - 32768&' % body
                r = self.result(I64 if op in (0x32, 0x33) else I32)
                self.emit('a = %s: %s = %s' % (ax, r, body)); return
            if op in (0x2E, 0x32):
                e = '((%s Xor 32768&) - 32768&)' % e
            return P('CLngLng(%s)' % e, I64, bits=17) if op in (0x32, 0x33) else P(e)
        if op == 0x29:
            if align >= 2 and not (x is None and ca % 4):
                self.tmp('a')
                r = self.result(I64)
                self.emit('a = %s: %s = (CLngLng(M(a + 1)) * 4294967296^) Or (CLngLng(M(a)) And 4294967295^)' % (self.widx(x, ca), r)); return
            return P('LD64(%s)' % ax, I64)
        if op == 0x34:
            return P('CLngLng(LD32(%s))' % ax, I64, bits=32)
        if op == 0x35:
            return P('(CLngLng(LD32(%s)) And 4294967295^)' % ax, I64, bits=33)
        if op == 0x2A:
            return P('BitsF32(LD32(%s))' % ax, F32)
        if op == 0x2B:
            return P('BitsF64(LD64(%s))' % ax, F64)
        raise Unsupported('load op %#x' % op)

    def store(self, op, align, off):
        v = self.st.pop(); base = self.st.pop()
        self.flush()
        x, ca = self.addr(base, off)
        ax = x if x is not None else lit32(ca)
        vx = self.val(v)
        if op == 0x38:
            op = 0x36; vx = 'F32Bits(%s)' % vx
        if op == 0x39:
            op = 0x37; vx = 'F64Bits(%s)' % vx
        if op in (0x3C, 0x3D, 0x3E):
            vx = W32 % vx; op = {0x3C: 0x3A, 0x3D: 0x3B, 0x3E: 0x36}[op]
        if op == 0x36:
            if align >= 2 and not (x is None and ca % 4):
                self.emit('M(%s) = %s' % (self.widx(x, ca), vx))
            else:
                self.emit('ST32 %s, %s' % (ax, vx))
            return
        if op == 0x37:
            if align >= 2 and not (x is None and ca % 4):
                self.tmp('a'); self.tmp('q', 'LongLong')
                self.emit('q = %s: a = %s: M(a) = %s: M(a + 1) = CLng((q And (-4294967296^)) \\ 4294967296^)' % (vx, self.widx(x, ca), W32 % 'q'))
            else:
                self.emit('ST64 %s, %s' % (ax, vx))
            return
        if op == 0x3A:
            self.tmp('a'); self.tmp('i')
            self.emit('a = %s: i = a \\ 4: M(i) = (M(i) And NM(a And 3)) Or BT((a And 3) * 256 + (%s And 255))' % (ax, vx)); return
        if op == 0x3B:
            if align >= 1 and not self.T.checked and x is not None:
                self.tmp('a'); self.tmp('i')
                self.emit('a = %s: i = a \\ 4: M(i) = (M(i) And NH(a And 2)) Or %s' % (ax, W32 % ('(CLngLng(%s And 65535&) * DV(a And 2))' % vx))); return
            self.emit('%s %s, %s' % ('ST16A' if (align >= 1 and self.T.checked) else 'ST16', ax, vx)); return
        raise Unsupported('store op %#x' % op)

    # ------------------------------------------------------------- arithmetic
    def binop32(self, op):
        b = self.st.pop(); a = self.st.pop()
        deps = a.deps | b.deps; mem = a.mem or b.mem
        ax, bx = self.val(a), self.val(b)
        def P(x, kind='v', c=None):
            self.push(x, I32, deps, mem, kind, c)
        if a.c is not None and b.c is not None and op in (0x6A, 0x6B, 0x6C, 0x71, 0x72, 0x73, 0x74):
            ua, ub = a.c & 0xFFFFFFFF, b.c & 0xFFFFFFFF
            r = {0x6A: ua + ub, 0x6B: ua - ub, 0x6C: ua * ub, 0x71: ua & ub, 0x72: ua | ub, 0x73: ua ^ ub,
                 0x74: ua << (ub & 31)}[op]
            return P(lit32(r), c=s32(r))
        if op == 0x6A:
            if b.c == 0: return self.st.append(a)
            if a.c == 0: return self.st.append(b)
            return P(W32 % ('(CLngLng(%s) + %s)' % (ax, bx)))
        if op == 0x6B:
            if b.c == 0: return self.st.append(a)
            return P(W32 % ('(CLngLng(%s) - %s)' % (ax, bx)))
        if op == 0x6C:
            return P(W32 % ('(CLngLng(%s) * %s)' % (ax, bx)))
        if op == 0x6D: return P('(%s \\ %s)' % (ax, bx))
        if op == 0x6F: return P('(%s Mod %s)' % (ax, bx))
        if op == 0x6E: return P('DivU32(%s, %s)' % (ax, bx))
        if op == 0x70: return P('RemU32(%s, %s)' % (ax, bx))
        if op == 0x71: return P('(%s And %s)' % (ax, bx))
        if op == 0x72: return P('(%s Or %s)' % (ax, bx))
        if op == 0x73: return P('(%s Xor %s)' % (ax, bx))
        if op == 0x74:
            if b.c is not None:
                k = b.c & 31
                if k == 0: return self.st.append(a)
                return P(W32 % ('(CLngLng(%s) * %d^)' % (ax, 1 << k)))
            return P(W32 % ('(CLngLng(%s) * P2(%s And 31))' % (ax, bx)))
        if op == 0x75:
            if b.c is not None:
                k = b.c & 31
                if k == 0: return self.st.append(a)
                if k == 31: return P('CLng(%s < 0)' % ax)
                return P('((%s And %s) \\ %d&)' % (ax, lit32(~((1 << k) - 1)), 1 << k))
            return P('ShrS32(%s, %s)' % (ax, bx))
        if op == 0x76:
            if b.c is not None:
                k = b.c & 31
                if k == 0: return self.st.append(a)
                if k == 31: return P('((%s < 0) And 1&)' % ax)
                if a.simple():
                    return P('(((%s And &H7FFFFFFF) \\ %d&) Or ((%s < 0) And %d&))' % (ax, 1 << k, ax, 1 << (31 - k)))
                return P('CLng((CLngLng(%s) And 4294967295^) \\ %d^)' % (ax, 1 << k))
            return P('ShrU32(%s, %s)' % (ax, bx))
        if op == 0x77: return P('Rotl32(%s, %s)' % (ax, bx))
        if op == 0x78: return P('Rotr32(%s, %s)' % (ax, bx))
        if op in CMP_S:
            return P('(%s %s %s)' % (ax, CMP_S[op], bx), 'b')
        if op in CMP_U:
            def u(e, x):
                if e.c is not None:
                    return lit32(e.c ^ 0x80000000)
                return '(%s Xor &H80000000)' % x
            return P('(%s %s %s)' % (u(a, ax), CMP_U[op], u(b, bx)), 'b')
        raise Unsupported('i32 op %#x' % op)

    def binop64(self, op):
        b = self.st.pop(); a = self.st.pop()
        deps = a.deps | b.deps; mem = a.mem or b.mem
        ax, bx = self.val(a), self.val(b)
        ab = a.bits if a.bits is not None else 64
        bb = b.bits if b.bits is not None else 64
        def P(x, t=I64, kind='v', bits=None):
            self.push(x, t, deps, mem, kind, bits=bits)
        if op == 0x7C:
            n = max(ab, bb) + 1
            return P('(%s + %s)' % (ax, bx), bits=n) if n <= 64 else P('Add64(%s, %s)' % (ax, bx))
        if op == 0x7D:
            n = max(ab, bb) + 1
            return P('(%s - %s)' % (ax, bx), bits=n) if n <= 64 else P('Sub64(%s, %s)' % (ax, bx))
        if op == 0x7E:
            n = ab + bb
            return P('(%s * %s)' % (ax, bx), bits=n) if n <= 64 else P('Mul64(%s, %s)' % (ax, bx))
        if op == 0x7F: return P('(%s \\ %s)' % (ax, bx))
        if op == 0x81: return P('(%s Mod %s)' % (ax, bx))
        if op == 0x80: return P('DivU64(%s, %s)' % (ax, bx))
        if op == 0x82: return P('RemU64(%s, %s)' % (ax, bx))
        if op == 0x83:
            n = None
            if b.c is not None and b.c >= 0: n = b.c.bit_length() + 1
            return P('(%s And %s)' % (ax, bx), bits=n)
        if op == 0x84: return P('(%s Or %s)' % (ax, bx), bits=(max(ab, bb) if max(ab, bb) < 64 else None))
        if op == 0x85: return P('(%s Xor %s)' % (ax, bx), bits=(max(ab, bb) if max(ab, bb) < 64 else None))
        if op == 0x86:
            if b.c is not None:
                k = b.c & 63
                if k == 0: return self.st.append(a)
                if ab + k <= 64:
                    return P('(%s * %d^)' % (ax, 1 << k), bits=ab + k)
                return P('Shl64(%s, %d&)' % (ax, k))
            return P('Shl64(%s, CLng(%s And 63^))' % (ax, bx))
        if op == 0x87:
            if b.c is not None:
                k = b.c & 63
                if k == 0: return self.st.append(a)
                if k == 63: return P('CLngLng(%s < 0)' % ax, bits=2)
                return P('((%s And %s) \\ %d^)' % (ax, lit64(~((1 << k) - 1)), 1 << k), bits=max(ab - k, 2))
            return P('ShrS64(%s, CLng(%s And 63^))' % (ax, bx))
        if op == 0x88:
            if b.c is not None:
                k = b.c & 63
                if k == 0: return self.st.append(a)
                nxt = self.ins[self.i + 1][0] if self.i + 1 < len(self.ins) else None
                if nxt == 0xA7 and k <= 32:
                    return P('((%s And %s) \\ %d^)' % (ax, lit64(~((1 << k) - 1)), 1 << k))
                return P('ShrU64(%s, %d&)' % (ax, k), bits=65 - k)
            return P('ShrU64(%s, CLng(%s And 63^))' % (ax, bx))
        if op == 0x89: return P('Rotl64(%s, CLng(%s And 63^))' % (ax, bx))
        if op == 0x8A: return P('Rotl64(%s, CLng((64^ - (%s And 63^)) And 63^))' % (ax, bx))
        if op in CMP64_S:
            return P('(%s %s %s)' % (ax, CMP64_S[op], bx), I32, 'b')
        if op in CMP64_U:
            return P('((%s Xor P2(63)) %s (%s Xor P2(63)))' % (ax, CMP64_U[op], bx), I32, 'b')
        raise Unsupported('i64 op %#x' % op)

    def floatop(self, op):
        if 0x5B <= op <= 0x66:
            b = self.st.pop(); a = self.st.pop()
            k = (op - 0x5B) % 6
            return self.push('(%s %s %s)' % (a.x, FCMP[k], b.x), I32, a.deps | b.deps, a.mem or b.mem, 'b')
        t = F32 if op < 0x99 else F64
        k = op - (0x8B if t == F32 else 0x99)
        wrap = 'CSng(%s)' if t == F32 else '%s'
        if k in FUN1:
            a = self.st.pop()
            return self.push(wrap % (FUN1[k] % a.x), t, a.deps, a.mem)
        b = self.st.pop(); a = self.st.pop()
        self.push(wrap % (FBIN[k] % (a.x, b.x)), t, a.deps | b.deps, a.mem or b.mem)

    def convert(self, op):
        e = self.st.pop(); x = self.val(e)
        P = lambda s, t, **kw: self.push(s, t, e.deps, e.mem, **kw)
        if op == 0xA7: return P(W32 % x, I32)
        if op in (0xA8, 0xAA, 0xFC00, 0xFC02): return P('TruncL(%s)' % x, I32)
        if op in (0xA9, 0xAB, 0xFC01, 0xFC03): return P('TruncUL(%s)' % x, I32)
        if op == 0xAC: return P('CLngLng(%s)' % x, I64, bits=32)
        if op == 0xAD: return P('(CLngLng(%s) And 4294967295^)' % x, I64, bits=33)
        if op in (0xAE, 0xB0, 0xFC04, 0xFC06): return P('TruncQ(%s)' % x, I64)
        if op in (0xAF, 0xB1, 0xFC05, 0xFC07): return P('TruncUQ(%s)' % x, I64)
        if op == 0xB2: return P('CSng(%s)' % x, F32)
        if op == 0xB3: return P('CSng(CLngLng(%s) And 4294967295^)' % x, F32)
        if op == 0xB4: return P('CSng(%s)' % x, F32)
        if op == 0xB5: return P('CSng(CvtUQ(%s))' % x, F32)
        if op == 0xB6: return P('CSng(%s)' % x, F32)
        if op == 0xB7: return P('CDbl(%s)' % x, F64)
        if op == 0xB8: return P('CDbl(CLngLng(%s) And 4294967295^)' % x, F64)
        if op == 0xB9: return P('CDbl(%s)' % x, F64)
        if op == 0xBA: return P('CvtUQ(%s)' % x, F64)
        if op == 0xBB: return P('CDbl(%s)' % x, F64)
        if op == 0xBC: return P('F32Bits(%s)' % x, I32)
        if op == 0xBD: return P('F64Bits(%s)' % x, I64)
        if op == 0xBE: return P('BitsF32(%s)' % x, F32)
        if op == 0xBF: return P('BitsF64(%s)' % x, F64)
        if op == 0xC0: return P('(((%s And 255&) Xor 128&) - 128&)' % x, I32)
        if op == 0xC1: return P('(((%s And 65535&) Xor 32768&) - 32768&)' % x, I32)
        if op == 0xC2: return P('(((%s And 255^) Xor 128^) - 128^)' % x, I64, bits=9)
        if op == 0xC3: return P('(((%s And 65535^) Xor 32768^) - 32768^)' % x, I64, bits=17)
        if op == 0xC4: return P('(((%s And 4294967295^) Xor 2147483648^) - 2147483648^)' % x, I64, bits=32)
        raise Unsupported('conversion %#x' % op)

    # ----------------------------------------------------------------- control
    def label(self, tgt):
        if tgt['kind'] == 'loop':
            return 'L%d' % tgt['start']
        if tgt['kind'] == 'func':
            return None
        return 'E%d' % tgt['start']

    def ret_stmts(self, x):
        if self.results:
            return ['%s = %s' % (self.name, x), 'Exit Function']
        return ['Exit Sub']

    def branch(self, tgt, conditional_expr=None):
        lab = self.label(tgt)
        if lab is None:
            body = self.ret_stmts(self.val(self.st[-1]) if self.results else None)
        else:
            body = []
            if tgt['arity']:
                dest = self.canon(tgt['height'], tgt['arity'][0])
                src = self.val(self.st[-1])
                if src != dest:
                    body.append('%s = %s' % (dest, src))
            body.append('GoTo %s' % lab)
        if conditional_expr is None:
            for s in body:
                self.emit(s)
        elif len(body) == 1:
            self.emit('If %s Then %s' % (conditional_expr, body[0]))
        else:
            self.emit('If %s Then' % conditional_expr)
            for s in body:
                self.emit(s)
            self.emit('End If')

    def translate(self):
        ins, ann = self.ins, self.ann
        for i, x in enumerate(ins):
            op = x[0]; a = ann[i]; self.i = i
            if a['dead']:
                if a.get('revive') and a['start'] is not None:
                    fr = self.frames[a['start']]
                    if op == ELSE:
                        self.emit('X%d:' % a['start']); self.reset(fr['types'])
                    else:
                        self.emit('E%d:' % a['start']); self.reset(fr['types'] + list(fr['results']))
                continue
            assert a['h'] == len(self.st), (self.name, i, a['h'], len(self.st))
            if op in (BLOCK, LOOP, IF):
                c = self.st.pop() if op == IF else None
                self.flush()
                self.frames[i] = dict(types=[e.t for e in self.st], results=x[1])
                if op == LOOP:
                    self.emit('L%d:' % i)
                elif op == IF:
                    tgt = 'X%d' % i if a['else'] is not None else 'E%d' % i
                    if c.kind == 'b':
                        self.emit('If Not %s Then GoTo %s' % (c.x, tgt))
                    else:
                        self.emit('If %s = 0 Then GoTo %s' % (c.x, tgt))
            elif op == ELSE:
                s = a['start']; fr = self.frames[s]
                self.flush(); self.emit('GoTo E%d' % s); self.emit('X%d:' % s); self.reset(fr['types'])
            elif op == END:
                s = a['start']
                if s is None:
                    if self.results:
                        self.emit('%s = %s' % (self.name, self.val(self.st[-1])))
                else:
                    self.flush()
                    if ins[s][0] != LOOP:
                        self.emit('E%d:' % s)
            elif op == 0x0C:
                tgt = a['br']
                if (tgt['arity'] and self.label(tgt) is not None) or (self.label(tgt) is None and self.results):
                    top = self.st.pop(); self.flush(); self.st.append(top)
                else:
                    self.flush()
                self.branch(tgt)
            elif op == 0x0D:
                c = self.st.pop(); self.flush()
                self.branch(a['br'], self.cond(c))
            elif op == 0x0E:
                idx = self.st.pop(); self.flush()
                self.emit('Select Case %s' % self.val(idx))
                for n, tgt in enumerate(a['brs']):
                    self.emit('Case %d' % n)
                    self.branch(tgt)
                self.emit('Case Else')
                self.branch(a['brd'])
                self.emit('End Select')
            elif op == 0x0F:
                for s in self.ret_stmts(self.val(self.st[-1]) if self.results else None):
                    self.emit(s)
            elif op == 0x00:
                self.emit('Trap')
            elif op == 0x01:
                pass
            elif op == 0x10 or op == 0x11:
                if op == 0x10:
                    pt, rt = func_type(self.m, x[1])
                    callee = self.T.fname(x[1])
                else:
                    pt, rt = self.m.types[x[1]]
                    callee = 'ci%d' % self.T.sig_id(pt, rt)
                    idx = self.st.pop()
                vals = [self.val(self.st.pop()) for _ in pt][::-1]
                args = ([self.val(idx)] if op == 0x11 else []) + vals
                self.flush()
                if rt:
                    r = self.canon(len(self.st), rt[0]); self.st.append(E(r, rt[0]))
                    self.emit('%s = %s(%s)' % (r, callee, ', '.join(args)))
                else:
                    self.emit('%s %s' % (callee, ', '.join(args)) if args else callee)
            elif op == 0x1A:
                self.st.pop()
            elif op == 0x1B:
                c = self.st.pop(); b = self.st.pop(); aa = self.st.pop()
                av, bv, cv = self.val(aa), self.val(b), self.cond(c)
                r = self.result(aa.t)
                self.emit('If %s Then %s = %s Else %s = %s' % (cv, r, av, r, bv))
            elif op == 0x20:
                n = 'l%d' % x[1]
                self.push(n, self.ltypes[x[1]], frozenset([n]))
            elif op == 0x21 or op == 0x22:
                n = 'l%d' % x[1]
                e = self.st.pop()
                self.hazard_local(n)
                if self.val(e) != n:
                    self.emit('%s = %s' % (n, self.val(e)))
                if op == 0x22:
                    self.push(n, self.ltypes[x[1]], frozenset([n]), bits=e.bits if False else None)
            elif op == 0x23:
                self.push('g%d' % x[1], self.m.globals[x[1]][0], frozenset(['G']))
            elif op == 0x24:
                e = self.st.pop(); self.hazard_mem()
                self.emit('g%d = %s' % (x[1], self.val(e)))
            elif op in wasm.LOAD_T:
                self.load(op, x[1], x[2])
            elif op in wasm.STORE_T:
                self.store(op, x[1], x[2])
            elif op == 0x3F:
                self.push('MemPages', I32, frozenset(['G']))
            elif op == 0x40:
                e = self.st.pop(); v = self.val(e)
                r = self.result(I32)
                self.emit('%s = MemGrow(%s)' % (r, v))
            elif op == 0xFC0A or op == 0xFC0B:
                n = self.val(self.st.pop()); s = self.val(self.st.pop()); d = self.val(self.st.pop())
                self.flush()
                self.emit('%s %s, %s, %s' % ('MemCopy' if op == 0xFC0A else 'MemFill', d, s, n))
            elif op == 0x41:
                self.push(lit32(x[1]), I32, c=s32(x[1]))
            elif op == 0x42:
                v = s64(x[1])
                self.push(lit64(v), I64, c=v, bits=v.bit_length() + 1)
            elif op == 0x43:
                self.push('BitsF32(%s)' % lit32(struct.unpack('<i', struct.pack('<f', x[1]))[0]), F32)
            elif op == 0x44:
                self.push('BitsF64(%s)' % lit64(struct.unpack('<q', struct.pack('<d', x[1]))[0]), F64)
            elif op == 0x45:
                e = self.st.pop()
                self.push('(Not %s)' % e.x if e.kind == 'b' else '(%s = 0)' % e.x, I32, e.deps, e.mem, 'b')
            elif op == 0x50:
                e = self.st.pop(); self.push('(%s = 0)' % e.x, I32, e.deps, e.mem, 'b')
            elif op in (0x67, 0x68, 0x69):
                e = self.st.pop(); self.push('%s(%s)' % ({0x67: 'Clz32', 0x68: 'Ctz32', 0x69: 'Popcnt32'}[op], self.val(e)), I32, e.deps, e.mem)
            elif op in (0x79, 0x7A, 0x7B):
                e = self.st.pop(); self.push('%s(%s)' % ({0x79: 'Clz64', 0x7A: 'Ctz64', 0x7B: 'Popcnt64'}[op], self.val(e)), I64, e.deps, e.mem, bits=8)
            elif 0x46 <= op <= 0x4F or 0x6A <= op <= 0x78:
                self.binop32(op)
            elif 0x51 <= op <= 0x5A or 0x7C <= op <= 0x8A:
                self.binop64(op)
            elif 0x5B <= op <= 0x66 or 0x8B <= op <= 0xA6:
                self.floatop(op)
            elif 0xA7 <= op <= 0xC4 or 0xFC00 <= op <= 0xFC07:
                self.convert(op)
            else:
                raise Unsupported(wasm.opclass(op) + ' (%#x)' % op)
        return self.finish()

    def finish(self):
        body = self.out
        refs = set(re.findall(r'GoTo ([A-Z]\d+)', '\n'.join(body)))
        body = [s for s in body if not (re.fullmatch(r'[LEX]\d+:', s) and s[:-1] not in refs)]
        pt, rt = self.m.types[self.m.func_types[self.k]]
        params = ', '.join('ByVal l%d As %s' % (i, VT[t]) for i, t in enumerate(pt))
        head = 'Public %s %s(%s)%s' % ('Function' if rt else 'Sub', self.name, params, ' As ' + VT[rt[0]] if rt else '')
        dims = ['l%d As %s' % (i, VT[t]) for i, t in enumerate(self.ltypes) if i >= self.nparams]
        dims += ['%s As %s' % (n, VT[t]) for n, t in sorted(self.used)]
        dims += ['%s As %s' % (n, t) for n, t in sorted(self.temps)]
        lines = [head]
        for i in range(0, len(dims), 8):
            lines.append('Dim ' + ', '.join(dims[i:i + 8]))
        lines += body
        lines.append('End Function' if rt else 'End Sub')
        return lines

class Translator:
    def __init__(self, path, checked=False, mem_pages=512):
        self.m = wasm.parse(path); self.checked = checked; self.mem_pages = mem_pages
        self.sigs = []

    def fname(self, idx):
        if idx < self.m.nimports:
            return 'imp_' + self.m.imports[idx][1]
        return 'f%d' % idx

    def sig_id(self, pt, rt):
        if (pt, rt) not in self.sigs:
            self.sigs.append((pt, rt))
        return self.sigs.index((pt, rt))

    def inline_small(self, fns):
        """Replace calls to one-line functions (e.g. fixed-point multiply) with their expression.
        A VBA call costs ~0.2 us, more than the arithmetic inside these."""
        m = self.m
        table = {}
        for k, f in fns:
            if len(f) != 3 or not f[0].startswith('Public Function '):
                continue
            name = 'f%d' % (k + m.nimports)
            hm = re.match(r'Public Function %s\((.*)\) As (\w+)$' % name, f[0])
            bm = re.match(r'%s = (.*)$' % name, f[1])
            if not hm or not bm:
                continue
            params = [p.split()[1] for p in hm.group(1).split(', ') if p]
            expr = bm.group(1)
            if len(expr) > 220 or re.search(r'\b(f\d+|ci\d+|imp_\w+)\(', expr):
                continue
            if any(len(re.findall(r'\b%s\b' % p, expr)) != 1 for p in params):
                continue
            table[name] = (params, expr)
        self.inlined = dict((n, 0) for n in table)
        def split_args(a):
            out = []; depth = 0; cur = ''
            for ch in a:
                if ch == '(': depth += 1
                elif ch == ')': depth -= 1
                if ch == ',' and depth == 0:
                    out.append(cur.strip()); cur = ''
                else:
                    cur += ch
            if cur.strip(): out.append(cur.strip())
            return out
        call = re.compile(r'^(\w+) = (f\d+)\((.*)\)$')
        res = []
        for k, f in fns:
            g = []
            for line in f:
                cm = call.match(line)
                if cm and cm.group(2) in table and line.count('(') == line.count(')'):
                    params, expr = table[cm.group(2)]
                    args = split_args(cm.group(3))
                    if len(args) == len(params) and len(line) + len(expr) < 700:
                        e = expr
                        # two-step substitution so an argument that mentions another parameter name is not re-substituted
                        for i, p in enumerate(params):
                            e = re.sub(r'\b%s\b' % p, '\x00%d\x00' % i, e)
                        for i, a in enumerate(args):
                            e = e.replace('\x00%d\x00' % i, '(' + a + ')' if not re.fullmatch(r'[\w&^]+', a) else a)
                        g.append('%s = %s' % (cm.group(1), e)); self.inlined[cm.group(2)] += 1
                        continue
                g.append(line)
            res.append((k, g))
        return res

    def data_init(self):
        m = self.m
        words = {}
        for off, data in m.datas:
            for j, byte in enumerate(data):
                a = off + j
                words[a // 4] = words.get(a // 4, 0) | (byte << (8 * (a % 4)))
        stm = ['M(%d) = %s' % (w, lit32(words[w])) for w in sorted(words) if words[w]]
        subs = []; names = []
        for i in range(0, len(stm), 400):
            n = 'DataInit%d' % (i // 400); names.append(n)
            subs.append(['Public Sub %s()' % n] + stm[i:i + 400] + ['End Sub'])
        return subs, names

    def build(self, lines_per_module=2500):
        m = self.m
        fns = []
        for k in range(len(m.codes)):
            try:
                fns.append((k, Fn(self, k).translate()))
            except Unsupported as e:
                raise Unsupported('%s in %s' % (e, m.names.get(k + m.nimports, k)))
        fns = self.inline_small(fns)
        rt = ['Option Explicit', 'Public MemPages As Long']
        for gi, (t, mut, v) in enumerate(m.globals):
            rt.append('Public g%d As %s' % (gi, VT[t]))
        table = {}
        for off, fs in m.elems:
            for j, f in enumerate(fs):
                table[off + j] = f
        for si, (pt, rt_) in enumerate(self.sigs):
            params = ''.join(', ByVal a%d As %s' % (i, VT[t]) for i, t in enumerate(pt))
            nm = 'ci%d' % si
            rt.append('Public %s %s(ByVal i As Long%s)%s' % ('Function' if rt_ else 'Sub', nm, params, ' As ' + VT[rt_[0]] if rt_ else ''))
            rt.append('Select Case i')
            for slot, f in sorted(table.items()):
                if func_type(m, f) == (pt, rt_):
                    args = ', '.join('a%d' % i for i in range(len(pt)))
                    call = '%s(%s)' % (self.fname(f), args) if rt_ else ('%s %s' % (self.fname(f), args)).strip()
                    rt.append('Case %d' % slot)
                    rt.append(('%s = ' % nm if rt_ else '') + call)
            rt.append('Case Else')
            rt.append('Trap')
            rt.append('End Select')
            rt.append('End Function' if rt_ else 'End Sub')
        for name, (kind, idx) in m.exports.items():
            if kind == 0:
                pt, rt_ = func_type(m, idx)
                params = ', '.join('ByVal a%d As %s' % (i, VT[t]) for i, t in enumerate(pt))
                args = ', '.join('a%d' % i for i in range(len(pt)))
                if rt_:
                    rt += ['Public Function x_%s(%s) As %s' % (name, params, VT[rt_[0]]), 'x_%s = f%d(%s)' % (name, idx, args), 'End Function']
                else:
                    rt += ['Public Sub x_%s(%s)' % (name, params), ('f%d %s' % (idx, args)).strip(), 'End Sub']
        subs, names = self.data_init()
        mods = []
        # data modules
        cur = ['Option Explicit']; dn = 0
        for sub in subs:
            if len(cur) + len(sub) > 6000 and len(cur) > 1:
                mods.append(('D%02d' % dn, '\n'.join(cur) + '\n')); dn += 1; cur = ['Option Explicit']
            cur += sub
        mods.append(('D%02d' % dn, '\n'.join(cur) + '\n'))
        rt.append('Public Sub ProgInit()')
        rt.append('MemPages = %d' % m.mem[0])
        for gi, (t, mut, v) in enumerate(m.globals):
            rt.append('g%d = %s' % (gi, lit32(v) if t == I32 else lit64(v)))
        rt += names
        rt.append('End Sub')
        mods.insert(0, ('Prog', '\n'.join(rt) + '\n'))
        cur = ['Option Explicit']; n = 0; touch = []
        def close():
            nonlocal cur, n
            nm = 'W%03d' % n
            cur += ['Public Sub Touch_%s()' % nm, 'End Sub']
            touch.append('Touch_%s' % nm)
            mods.append((nm, '\n'.join(cur) + '\n')); n += 1; cur = ['Option Explicit']
        for k, f in fns:
            if len(cur) + len(f) > lines_per_module and len(cur) > 1:
                close()
            cur.append("' %s" % m.names.get(k + m.nimports, ''))
            cur += f
        close()
        mods.append(('Touch', 'Option Explicit\nPublic Sub TouchAll()\n' + '\n'.join(touch) + '\nEnd Sub\n'))
        sizes = sorted(((len(f), sum(len(s) for s in f), max(len(s) for s in f), m.names.get(k + m.nimports, str(k))) for k, f in fns), reverse=True)
        stats = dict(functions=len(fns), wasm_instrs=sum(len(c[1]) for c in m.codes), vba_lines=sum(len(f) for _, f in fns),
                     modules=len(mods), largest=sizes[:6], longest_line=max(s[2] for s in sizes))
        return mods, stats
