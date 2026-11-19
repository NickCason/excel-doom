"""Spike translator: WebAssembly -> VBA source (static recompilation).

Memory model is selectable: 'L' = linear memory as a Long array (word-addressed),
'B' = Byte array. Integer-only; i64 arithmetic is unwrapped LongLong (good enough
for Doom's fixed-point maths, not general)."""
import wasm
from wasm import I32, I64, BLOCK, LOOP, IF, ELSE, END, func_type

VT = {I32: 'Long', I64: 'LongLong'}
W32 = 'CLng(((%s And 4294967295^) Xor 2147483648^) - 2147483648^)'

def s32(v):
    v &= 0xFFFFFFFF
    return v - (1 << 32) if v >= 0x80000000 else v

def lit32(v):
    v = s32(v)
    if v == -2147483648:
        return '&H80000000'
    return '%d&' % v if v >= 0 else '(%d&)' % v

def lit64(v):
    v &= (1 << 64) - 1
    if v >= 1 << 63:
        v -= 1 << 64
    if v == -(1 << 63):
        return '(-9223372036854775807^ - 1^)'
    return '%d^' % v if v >= 0 else '(%d^)' % v

class E:
    __slots__ = ('x', 't', 'deps', 'mem', 'kind', 'c')
    def __init__(self, x, t, deps=frozenset(), mem=False, kind='v', c=None):
        self.x = x; self.t = t; self.deps = deps; self.mem = mem; self.kind = kind; self.c = c
    def simple(self):
        return self.c is not None or self.x.isidentifier()

CMP_S = {0x46: '=', 0x47: '<>', 0x48: '<', 0x4A: '>', 0x4C: '<=', 0x4E: '>='}
CMP_U = {0x49: '<', 0x4B: '>', 0x4D: '<=', 0x4F: '>='}
CMP64_S = {0x51: '=', 0x52: '<>', 0x53: '<', 0x55: '>', 0x57: '<=', 0x59: '>='}

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
        if t not in VT:
            raise Unsupported('float value')
        n = ('s%d' if t == I32 else 'j%d') % d
        self.used.add((n, t))
        return n

    def push(self, x, t, deps=frozenset(), mem=False, kind='v', c=None):
        self.st.append(E(x, t, deps, mem, kind, c))
        if len(x) > 200:
            self.flush()

    def val(self, e):
        return '((%s) And 1&)' % e.x if e.kind == 'b' else e.x

    def cond(self, e):
        if e.kind == 'b':
            return e.x
        return '(%s <> 0)' % e.x

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
        """Operands already popped: settle everything below, return the canonical slot."""
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

    def load(self, op, align, off):
        base = self.st.pop()
        x, ca = self.addr(base, off)
        ax = x if x is not None else lit32(ca)
        deps = base.deps
        mode = self.T.mode
        if op == 0x29:
            self.push('LD64(%s)' % ax, I64, deps, True); return
        if op not in (0x28, 0x2C, 0x2D, 0x2E, 0x2F):
            raise Unsupported('load op %#x' % op)
        if mode == 'B':
            if op in (0x2D, 0x2C):
                e = 'B(%s)' % ax
                if op == 0x2C:
                    e = '((%s Xor 128) - 128)' % e
                self.push('(%s)' % e if op == 0x2C else e, I32, deps, True); return
            a = self.tmp('a')
            r = None
            pre = '%s = %s: ' % (a, ax)
            if op == 0x28:
                body = 'B(a) Or B(a + 1) * 256& Or B(a + 2) * 65536 Or (B(a + 3) And 127) * 16777216 Or ((B(a + 3) > 127) And &H80000000)'
            else:
                body = 'B(a) Or B(a + 1) * 256&'
                if op == 0x2E:
                    body = '((%s) Xor 32768&) - 32768&' % body
            r = self.result(I32)
            self.emit('%s%s = %s' % (pre, r, body)); return
        # ---- Long-array memory
        if op == 0x28:
            if align >= 2:
                self.push('M(%d)' % (ca // 4) if x is None and ca % 4 == 0 else 'M(%s \\ 4)' % ax, I32, deps, True)
            else:
                self.push('LD32(%s)' % ax, I32, deps, True)
            return
        if op in (0x2D, 0x2C):
            if self.T.byte_helper:
                e = 'LD8(%s)' % ax
                if op == 0x2C:
                    e = '((%s Xor 128) - 128)' % e
                self.push(e, I32, deps, True); return
            if x is None:
                k = ca & 3
                e = '((M(%d) And %s) \\ %d&)' % (ca // 4, lit32(0xFF << (8 * k)), 1 << (8 * k))
                if k == 3:
                    e = '(%s And 255)' % e
                if op == 0x2C:
                    e = '((%s Xor 128) - 128)' % e
                self.push(e, I32, deps, True); return
            a = self.tmp('a')
            body = '((M(a \\ 4) And MK(a And 3)) \\ DV(a And 3)) And 255'
            if op == 0x2C:
                body = '((%s) Xor 128) - 128' % body
            r = self.result(I32)
            self.emit('a = %s: %s = %s' % (ax, r, body)); return
        # 16-bit
        if align >= 1 and x is not None:
            a = self.tmp('a')
            body = '((M(a \\ 4) And MH(a And 2)) \\ DV(a And 2)) And 65535&'
            if op == 0x2E:
                body = '((%s) Xor 32768&) - 32768&' % body
            r = self.result(I32)
            self.emit('a = %s: %s = %s' % (ax, r, body)); return
        e = 'LD16(%s)' % ax
        if op == 0x2E:
            e = '((%s Xor 32768&) - 32768&)' % e
        self.push(e, I32, deps, True)

    def store(self, op, align, off):
        v = self.st.pop(); base = self.st.pop()
        self.flush()
        x, ca = self.addr(base, off)
        ax = x if x is not None else lit32(ca)
        vx = self.val(v)
        mode = self.T.mode
        if op == 0x37:
            self.emit('ST64 %s, %s' % (ax, vx)); return
        if op not in (0x36, 0x3A, 0x3B):
            raise Unsupported('store op %#x' % op)
        if mode == 'B':
            if op == 0x3A:
                self.emit('B(%s) = %s And 255' % (ax, vx)); return
            self.tmp('a'); self.tmp('v')
            s = 'a = %s: v = %s: B(a) = v And 255: B(a + 1) = (v And 65280) \\ 256' % (ax, vx)
            if op == 0x36:
                s += ': B(a + 2) = (v And 16711680) \\ 65536: B(a + 3) = ((v And &H7F000000) \\ 16777216) Or ((v < 0) And 128)'
            self.emit(s); return
        if op == 0x36:
            if align >= 2:
                self.emit('%s = %s' % ('M(%d)' % (ca // 4) if x is None and ca % 4 == 0 else 'M(%s \\ 4)' % ax, vx))
            else:
                self.emit('ST32 %s, %s' % (ax, vx))
            return
        if op == 0x3A:
            if self.T.byte_helper:
                self.emit('ST8 %s, %s' % (ax, vx)); return
            self.tmp('a'); self.tmp('i')
            self.emit('a = %s: i = a \\ 4: M(i) = (M(i) And NM(a And 3)) Or BT((a And 3) * 256 + (%s And 255))' % (ax, vx)); return
        self.emit('ST16 %s, %s' % (ax, vx))

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
        def P(x, t=I64, kind='v'):
            self.push(x, t, deps, mem, kind)
        if op == 0x7C: return P('(%s + %s)' % (ax, bx))
        if op == 0x7D: return P('(%s - %s)' % (ax, bx))
        if op == 0x7E: return P('(%s * %s)' % (ax, bx))
        if op == 0x7F: return P('(%s \\ %s)' % (ax, bx))
        if op == 0x81: return P('(%s Mod %s)' % (ax, bx))
        if op == 0x83: return P('(%s And %s)' % (ax, bx))
        if op == 0x84: return P('(%s Or %s)' % (ax, bx))
        if op == 0x85: return P('(%s Xor %s)' % (ax, bx))
        if op == 0x86 and b.c is not None:
            return P('(%s * %d^)' % (ax, 1 << (b.c & 63)))
        if op == 0x87 and b.c is not None:
            k = b.c & 63
            return P('((%s And %s) \\ %d^)' % (ax, lit64(~((1 << k) - 1)), 1 << k))
        if op == 0x88:
            if b.c is not None:
                k = b.c & 63
                nxt = self.ins[self.i + 1][0] if self.i + 1 < len(self.ins) else None
                if nxt == 0xA7 and k <= 32:
                    # only the low 32 bits survive the following wrap, so a floor shift is equivalent
                    return P('((%s And %s) \\ %d^)' % (ax, lit64(~((1 << k) - 1)), 1 << k))
                return P('ShrU64(%s, %d&)' % (ax, k))
            return P('ShrU64(%s, CLng(%s And 63^))' % (ax, bx))
        if op in CMP64_S:
            return P('(%s %s %s)' % (ax, CMP64_S[op], bx), I32, 'b')
        raise Unsupported('i64 op %#x' % op)

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
        """Emit a (possibly conditional) jump to tgt. Stack must already be flushed below top."""
        lab = self.label(tgt)
        moves = []
        if lab is None:
            body = self.ret_stmts(self.val(self.st[-1]) if self.results else None)
        else:
            if tgt['arity']:
                assert len(tgt['arity']) == 1
                dest = self.canon(tgt['height'], tgt['arity'][0])
                src = self.val(self.st[-1])
                if src != dest:
                    moves.append('%s = %s' % (dest, src))
            body = moves + ['GoTo %s' % lab]
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
                if a.get('revive'):
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
                if tgt['arity'] and self.label(tgt) is not None:
                    top = self.st.pop(); self.flush(); self.st.append(top)
                else:
                    if self.label(tgt) is None and self.results:
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
                    args = []
                else:
                    pt, rt = self.m.types[x[1]]
                    sig = self.T.sig_id(pt, rt)
                    callee = 'ci%d' % sig
                    args = [None]
                    idx = self.st.pop()
                vals = [self.val(self.st.pop()) for _ in pt][::-1]
                if op == 0x11:
                    args = [self.val(idx)]
                args += vals
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
                av, bv = self.val(aa), self.val(b)
                r = self.result(aa.t)
                self.emit('If %s Then %s = %s Else %s = %s' % (self.cond(c), r, av, r, bv))
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
                    self.push(n, self.ltypes[x[1]], frozenset([n]))
            elif op == 0x23:
                self.push('g%d' % x[1], self.m.globals[x[1]][0], frozenset(['G']))
            elif op == 0x24:
                e = self.st.pop(); self.hazard_mem()
                self.emit('g%d = %s' % (x[1], self.val(e)))
            elif op in wasm.LOAD_T:
                self.load(op, x[1], x[2])
            elif op in wasm.STORE_T:
                self.store(op, x[1], x[2])
            elif op == 0x41:
                self.push(lit32(x[1]), I32, c=s32(x[1]))
            elif op == 0x42:
                self.push(lit64(x[1]), I64, c=x[1])
            elif op == 0x45:
                e = self.st.pop()
                if e.kind == 'b':
                    self.push('(Not %s)' % e.x, I32, e.deps, e.mem, 'b')
                else:
                    self.push('(%s = 0)' % e.x, I32, e.deps, e.mem, 'b')
            elif op == 0x50:
                e = self.st.pop(); self.push('(%s = 0)' % e.x, I32, e.deps, e.mem, 'b')
            elif 0x46 <= op <= 0x4F or 0x6A <= op <= 0x78:
                self.binop32(op)
            elif 0x51 <= op <= 0x5A or 0x7C <= op <= 0x8A:
                self.binop64(op)
            elif op == 0xA7:
                e = self.st.pop(); self.push(W32 % e.x, I32, e.deps, e.mem)
            elif op == 0xAC:
                e = self.st.pop(); self.push('CLngLng(%s)' % self.val(e), I64, e.deps, e.mem)
            elif op == 0xAD:
                e = self.st.pop(); self.push('(CLngLng(%s) And 4294967295^)' % self.val(e), I64, e.deps, e.mem)
            elif op == 0xC0:
                e = self.st.pop(); self.push('(((%s And 255&) Xor 128&) - 128&)' % self.val(e), I32, e.deps, e.mem)
            elif op == 0xC1:
                e = self.st.pop(); self.push('(((%s And 65535&) Xor 32768&) - 32768&)' % self.val(e), I32, e.deps, e.mem)
            elif op in (0x67, 0x68, 0x69):
                e = self.st.pop(); self.push('%s(%s)' % ({0x67: 'Clz32', 0x68: 'Ctz32', 0x69: 'Popcnt32'}[op], self.val(e)), I32, e.deps, e.mem)
            else:
                raise Unsupported(wasm.opclass(op) + ' (%#x)' % op)
        return self.finish()

    def finish(self):
        body = self.out
        used = set()
        for s in body:
            for w in s.replace(':', ' ').split():
                pass
        import re
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

RUNTIME_DECL = r'''
Public MK(0 To 3) As Long, DV(0 To 3) As Long, NM(0 To 3) As Long, MH(0 To 2) As Long
Public BT(0 To 1023) As Long
Public P2(0 To 63) As LongLong
'''

RUNTIME = r'''
Public Sub Trap()
    Err.Raise vbObjectError + 77, , "wasm trap"
End Sub

Public Sub RtTables()
    Dim k As Long, v As Long
    MK(0) = &HFF&: MK(1) = &HFF00&: MK(2) = &HFF0000: MK(3) = &HFF000000
    DV(0) = 1: DV(1) = 256: DV(2) = 65536: DV(3) = 16777216
    For k = 0 To 3: NM(k) = Not MK(k): Next k
    MH(0) = &HFFFF&: MH(2) = &HFFFF0000
    For k = 0 To 2
        For v = 0 To 255: BT(k * 256 + v) = v * DV(k): Next v
    Next k
    For v = 0 To 255: BT(768 + v) = ((v And 127) * 16777216) Or ((v > 127) And &H80000000): Next v
    P2(0) = 1
    For k = 1 To 62: P2(k) = P2(k - 1) * 2: Next k
    P2(63) = -9223372036854775807^ - 1^
End Sub

Public Function W32(ByVal x As LongLong) As Long
    W32 = CLng(((x And 4294967295^) Xor 2147483648^) - 2147483648^)
End Function
Public Function ShrS32(ByVal a As Long, ByVal k As Long) As Long
    Dim x As LongLong
    k = k And 31: x = a
    ShrS32 = CLng((x - (x And (P2(k) - 1^))) \ P2(k))
End Function
Public Function ShrU32(ByVal a As Long, ByVal k As Long) As Long
    ShrU32 = W32((CLngLng(a) And 4294967295^) \ P2(k And 31))
End Function
Public Function DivU32(ByVal a As Long, ByVal b As Long) As Long
    DivU32 = W32((CLngLng(a) And 4294967295^) \ (CLngLng(b) And 4294967295^))
End Function
Public Function RemU32(ByVal a As Long, ByVal b As Long) As Long
    RemU32 = W32((CLngLng(a) And 4294967295^) Mod (CLngLng(b) And 4294967295^))
End Function
Public Function Rotl32(ByVal a As Long, ByVal k As Long) As Long
    Dim u As LongLong
    k = k And 31: u = CLngLng(a) And 4294967295^
    If k = 0 Then Rotl32 = a Else Rotl32 = W32((u * P2(k)) Or (u \ P2(32 - k)))
End Function
Public Function Rotr32(ByVal a As Long, ByVal k As Long) As Long
    Rotr32 = Rotl32(a, 32 - (k And 31))
End Function
Public Function ShrU64(ByVal v As LongLong, ByVal k As Long) As LongLong
    k = k And 63
    If k = 0 Then
        ShrU64 = v
    ElseIf v >= 0 Then
        ShrU64 = v \ P2(k)
    Else
        ShrU64 = ((v And 9223372036854775807^) \ P2(k)) Or P2(63 - k)
    End If
End Function
Public Function LD64(ByVal a As Long) As LongLong
    LD64 = (CLngLng(LD32(a + 4)) * 4294967296^) Or (CLngLng(LD32(a)) And 4294967295^)
End Function
Public Sub ST64(ByVal a As Long, ByVal v As LongLong)
    ST32 a, W32(v)
    ST32 a + 4, CLng((v And (-4294967296^)) \ 4294967296^)
End Sub
'''

RUNTIME_L_DECL = 'Public M() As Long'
RUNTIME_L = r'''
Public Function LD8(ByVal a As Long) As Long
    LD8 = ((M(a \ 4) And MK(a And 3)) \ DV(a And 3)) And 255
End Function
Public Function LD16(ByVal a As Long) As Long
    LD16 = LD8(a) Or LD8(a + 1) * 256&
End Function
Public Function LD32(ByVal a As Long) As Long
    If (a And 3) = 0 Then
        LD32 = M(a \ 4)
    Else
        LD32 = LD8(a) Or LD8(a + 1) * 256& Or LD8(a + 2) * 65536 Or BT(768 + LD8(a + 3))
    End If
End Function
Public Sub ST8(ByVal a As Long, ByVal v As Long)
    Dim i As Long
    i = a \ 4
    M(i) = (M(i) And NM(a And 3)) Or BT((a And 3) * 256 + (v And 255))
End Sub
Public Sub ST16(ByVal a As Long, ByVal v As Long)
    ST8 a, v: ST8 a + 1, (v And &HFF00&) \ 256
End Sub
Public Sub ST32(ByVal a As Long, ByVal v As Long)
    If (a And 3) = 0 Then
        M(a \ 4) = v
    Else
        ST8 a, v: ST8 a + 1, (v And &HFF00&) \ 256: ST8 a + 2, (v And &HFF0000) \ 65536
        ST8 a + 3, ((v And &HFF000000) \ 16777216) And 255
    End If
End Sub
'''

RUNTIME_B_DECL = 'Public B() As Byte'
RUNTIME_B = r'''
Public Function LD32(ByVal a As Long) As Long
    LD32 = B(a) Or B(a + 1) * 256& Or B(a + 2) * 65536 Or (B(a + 3) And 127) * 16777216 Or ((B(a + 3) > 127) And &H80000000)
End Function
Public Sub ST32(ByVal a As Long, ByVal v As Long)
    B(a) = v And 255: B(a + 1) = (v And 65280) \ 256: B(a + 2) = (v And 16711680) \ 65536
    B(a + 3) = ((v And &H7F000000) \ 16777216) Or ((v < 0) And 128)
End Sub
'''

class Translator:
    def __init__(self, path, mode='L', byte_helper=False):
        self.m = wasm.parse(path); self.mode = mode; self.byte_helper = byte_helper
        self.sigs = []

    def fname(self, idx):
        if idx < self.m.nimports:
            return 'imp_' + self.m.imports[idx][1]
        return 'f%d' % idx

    def sig_id(self, pt, rt):
        if (pt, rt) not in self.sigs:
            self.sigs.append((pt, rt))
        return self.sigs.index((pt, rt))

    def data_init(self):
        """VBA subs that load data segments into memory."""
        m = self.m
        img = {}
        for off, data in m.datas:
            for j, byte in enumerate(data):
                img[off + j] = byte
        stm = []
        if self.mode == 'L':
            words = {}
            for a, byte in img.items():
                words[a // 4] = words.get(a // 4, 0) | (byte << (8 * (a % 4)))
            for w in sorted(words):
                if words[w]:
                    stm.append('M(%d) = %s' % (w, lit32(words[w])))
        else:
            for a in sorted(img):
                if img[a]:
                    stm.append('B(%d) = %d' % (a, img[a]))
        subs = []; names = []
        for i in range(0, len(stm), 300):
            n = 'DataInit%d' % (i // 300); names.append(n)
            subs += ['Private Sub %s()' % n] + stm[i:i + 300] + ['End Sub']
        return subs, names

    def build(self, lines_per_module=3000):
        m = self.m
        fns = []; failed = []
        nins = 0
        for k in range(len(m.codes)):
            try:
                fns.append(Fn(self, k).translate()); nins += len(m.codes[k][1])
            except Unsupported as e:
                failed.append((k + m.nimports, str(e)))
        pages = m.mem[0]
        rt = ['Option Explicit', RUNTIME_DECL, RUNTIME_L_DECL if self.mode == 'L' else RUNTIME_B_DECL]
        for gi, (t, mut, v) in enumerate(m.globals):
            rt.append('Public g%d As %s' % (gi, VT[t]))
        rt += [RUNTIME, RUNTIME_L if self.mode == 'L' else RUNTIME_B]
        subs, names = self.data_init()
        rt += subs
        rt.append('Public Sub RtInit()')
        rt.append('RtTables')
        rt.append('ReDim M(0 To %d)' % (pages * 16384 - 1) if self.mode == 'L' else 'ReDim B(0 To %d)' % (pages * 65536 - 1))
        for gi, (t, mut, v) in enumerate(m.globals):
            rt.append('g%d = %s' % (gi, lit32(v) if t == I32 else lit64(v)))
        rt += names
        rt.append('End Sub')
        # indirect-call dispatchers
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
                    rt.append('Case %d: %s' % (slot, ('%s = ' % nm if rt_ else '') + call))
            rt.append('Case Else: Trap')
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
        mods = [('Rt', '\n'.join(rt) + '\n')]
        cur = ['Option Explicit']; n = 0
        for f in fns:
            if len(cur) + len(f) > lines_per_module and len(cur) > 1:
                mods.append(('W%02d' % n, '\n'.join(cur) + '\n')); n += 1; cur = ['Option Explicit']
            cur += f + ['']
        mods.append(('W%02d' % n, '\n'.join(cur) + '\n'))
        stats = dict(functions=len(fns), failed=failed, wasm_instrs=nins,
                     vba_lines=sum(len(f) for f in fns), max_fn_lines=max(len(f) for f in fns) if fns else 0)
        return mods, stats

if __name__ == '__main__':
    import sys
    T = Translator(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else 'L')
    mods, stats = T.build()
    stats['failed'] = stats['failed'][:8]
    print(stats)
    if len(sys.argv) > 3:
        open(sys.argv[3], 'w').write('\n'.join(src for _, src in mods))
