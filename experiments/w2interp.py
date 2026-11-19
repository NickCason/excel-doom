"""Spike emulator: pre-decode WebAssembly into flat arrays and generate a VBA
interpreter loop (On..GoTo dispatch) that runs it. Same memory model and helpers
as the translator's Long-array mode, so the two are directly comparable."""
import wasm
from wasm import I32, I64, BLOCK, LOOP, IF, ELSE, END, func_type
import w2vba

WR = 'CLng(((%s And 4294967295^) Xor 2147483648^) - 2147483648^)'
H = {}   # name -> handler statements (list of lines)
def h(name, *lines):
    H[name] = list(lines)

h('LGET', 'S(sp) = S(fp + IA(pc)): sp = sp + 1')
h('LGET64', 'SJ(sp) = SJ(fp + IA(pc)): sp = sp + 1')
h('LSET', 'sp = sp - 1: S(fp + IA(pc)) = S(sp)')
h('LSET64', 'sp = sp - 1: SJ(fp + IA(pc)) = SJ(sp)')
h('LTEE', 'S(fp + IA(pc)) = S(sp - 1)')
h('LTEE64', 'SJ(fp + IA(pc)) = SJ(sp - 1)')
h('GGET', 'S(sp) = G(IA(pc)): sp = sp + 1')
h('GSET', 'sp = sp - 1: G(IA(pc)) = S(sp)')
h('CONST', 'S(sp) = IA(pc): sp = sp + 1')
h('CONST64', 'SJ(sp) = KJ(IA(pc)): sp = sp + 1')
h('LOAD32', 'S(sp - 1) = M((S(sp - 1) + IA(pc)) \\ 4)')
h('LOAD32X', 'S(sp - 1) = LD32(S(sp - 1) + IA(pc))')
h('LOAD8U', 'a = S(sp - 1) + IA(pc): S(sp - 1) = ((M(a \\ 4) And MK(a And 3)) \\ DV(a And 3)) And 255')
h('LOAD8S', 'a = S(sp - 1) + IA(pc): S(sp - 1) = ((((M(a \\ 4) And MK(a And 3)) \\ DV(a And 3)) And 255) Xor 128) - 128')
h('LOAD16U', 'S(sp - 1) = LD16(S(sp - 1) + IA(pc))')
h('LOAD16S', 'S(sp - 1) = (LD16(S(sp - 1) + IA(pc)) Xor 32768&) - 32768&')
h('LOAD64', 'SJ(sp - 1) = LD64(S(sp - 1) + IA(pc))')
h('STORE32', 'sp = sp - 2: M((S(sp) + IA(pc)) \\ 4) = S(sp + 1)')
h('STORE32X', 'sp = sp - 2: ST32 S(sp) + IA(pc), S(sp + 1)')
h('STORE8', 'sp = sp - 2: a = S(sp) + IA(pc): i = a \\ 4: M(i) = (M(i) And NM(a And 3)) Or BT((a And 3) * 256 + (S(sp + 1) And 255))')
h('STORE16', 'sp = sp - 2: ST16 S(sp) + IA(pc), S(sp + 1)')
h('STORE64', 'sp = sp - 2: ST64 S(sp) + IA(pc), SJ(sp + 1)')
for nm, op in [('ADD', '+'), ('SUB', '-'), ('MUL', '*')]:
    h(nm, 'sp = sp - 1: S(sp - 1) = ' + WR % ('(CLngLng(S(sp - 1)) %s S(sp))' % op))
for nm, op in [('AND', 'And'), ('OR', 'Or'), ('XOR', 'Xor'), ('DIVS', '\\'), ('REMS', 'Mod')]:
    h(nm, 'sp = sp - 1: S(sp - 1) = S(sp - 1) %s S(sp)' % op)
h('DIVU', 'sp = sp - 1: S(sp - 1) = DivU32(S(sp - 1), S(sp))')
h('REMU', 'sp = sp - 1: S(sp - 1) = RemU32(S(sp - 1), S(sp))')
h('SHL', 'sp = sp - 1: S(sp - 1) = ' + WR % '(CLngLng(S(sp - 1)) * P2(S(sp) And 31))')
h('SHRS', 'sp = sp - 1: S(sp - 1) = ShrS32(S(sp - 1), S(sp))')
h('SHRU', 'sp = sp - 1: S(sp - 1) = ShrU32(S(sp - 1), S(sp))')
for nm, op in [('EQ', '='), ('NE', '<>'), ('LTS', '<'), ('GTS', '>'), ('LES', '<='), ('GES', '>=')]:
    h(nm, 'sp = sp - 1: S(sp - 1) = (S(sp - 1) %s S(sp)) And 1' % op)
    h(nm + '64', 'sp = sp - 1: S(sp - 1) = (SJ(sp - 1) %s SJ(sp)) And 1' % op)
for nm, op in [('LTU', '<'), ('GTU', '>'), ('LEU', '<='), ('GEU', '>=')]:
    h(nm, 'sp = sp - 1: S(sp - 1) = ((S(sp - 1) Xor &H80000000) %s (S(sp) Xor &H80000000)) And 1' % op)
h('EQZ', 'S(sp - 1) = (S(sp - 1) = 0) And 1')
h('EQZ64', 'S(sp - 1) = (SJ(sp - 1) = 0) And 1')
h('EXT8S', 'S(sp - 1) = ((S(sp - 1) And 255&) Xor 128&) - 128&')
h('EXT16S', 'S(sp - 1) = ((S(sp - 1) And 65535&) Xor 32768&) - 32768&')
for nm, op in [('ADD64', '+'), ('SUB64', '-'), ('MUL64', '*'), ('DIVS64', '\\'), ('AND64', 'And'), ('OR64', 'Or'), ('XOR64', 'Xor')]:
    h(nm, 'sp = sp - 1: SJ(sp - 1) = SJ(sp - 1) %s SJ(sp)' % op)
h('SHL64', 'sp = sp - 1: SJ(sp - 1) = SJ(sp - 1) * P2(CLng(SJ(sp) And 63))')
h('SHRS64', 'sp = sp - 1: x = P2(CLng(SJ(sp) And 63)): SJ(sp - 1) = (SJ(sp - 1) - (SJ(sp - 1) And (x - 1^))) \\ x')
h('SHRU64', 'sp = sp - 1: SJ(sp - 1) = ShrU64(SJ(sp - 1), CLng(SJ(sp) And 63))')
h('WRAP', 'S(sp - 1) = ' + WR % 'SJ(sp - 1)')
h('EXTS', 'SJ(sp - 1) = S(sp - 1)')
h('EXTU', 'SJ(sp - 1) = CLngLng(S(sp - 1)) And 4294967295^')
h('SELECT', 'sp = sp - 2: If S(sp + 1) = 0 Then S(sp - 1) = S(sp)')
h('SELECT64', 'sp = sp - 2: If S(sp + 1) = 0 Then SJ(sp - 1) = SJ(sp)')
h('DROP', 'sp = sp - 1')
h('BR', 'pc = IA(pc): GoTo DSP')
h('BRADJ0', 'sp = fp + IB(pc): pc = IA(pc): GoTo DSP')
h('BRADJ1', 'S(fp + IB(pc)) = S(sp - 1): SJ(fp + IB(pc)) = SJ(sp - 1): sp = fp + IB(pc) + 1: pc = IA(pc): GoTo DSP')
h('BRIF', 'sp = sp - 1: If S(sp) <> 0 Then pc = IA(pc): GoTo DSP')
h('BRIFADJ0', 'sp = sp - 1: If S(sp) <> 0 Then sp = fp + IB(pc): pc = IA(pc): GoTo DSP')
h('BRIFADJ1', 'sp = sp - 1: If S(sp) <> 0 Then S(fp + IB(pc)) = S(sp - 1): SJ(fp + IB(pc)) = SJ(sp - 1): sp = fp + IB(pc) + 1: pc = IA(pc): GoTo DSP')
h('BRZ', 'sp = sp - 1: If S(sp) = 0 Then pc = IA(pc): GoTo DSP')
h('BRT', 'sp = sp - 1: k = S(sp): If k < 0 Or k >= IB(pc) Then k = IB(pc)',
  'k = IA(pc) + k * 2: sp = fp + BP(k + 1): pc = BP(k): GoTo DSP')
h('CALL', 'f = IA(pc): GoTo DOCALL')
h('CALLI', 'sp = sp - 1: f = TBL(S(sp)): GoTo DOCALL')
h('RET0', 'sp = fp: GoTo DORET')
h('RET1', 'S(fp) = S(sp - 1): SJ(fp) = SJ(sp - 1): sp = fp + 1: GoTo DORET')
h('UNREACH', 'Trap')
OPNUM = {n: i + 1 for i, n in enumerate(H)}

BIN32 = {0x6A: 'ADD', 0x6B: 'SUB', 0x6C: 'MUL', 0x6D: 'DIVS', 0x6E: 'DIVU', 0x6F: 'REMS', 0x70: 'REMU', 0x71: 'AND',
         0x72: 'OR', 0x73: 'XOR', 0x74: 'SHL', 0x75: 'SHRS', 0x76: 'SHRU', 0x46: 'EQ', 0x47: 'NE', 0x48: 'LTS',
         0x49: 'LTU', 0x4A: 'GTS', 0x4B: 'GTU', 0x4C: 'LES', 0x4D: 'LEU', 0x4E: 'GES', 0x4F: 'GEU', 0x45: 'EQZ',
         0xC0: 'EXT8S', 0xC1: 'EXT16S'}
BIN64 = {0x7C: 'ADD64', 0x7D: 'SUB64', 0x7E: 'MUL64', 0x7F: 'DIVS64', 0x83: 'AND64', 0x84: 'OR64', 0x85: 'XOR64',
         0x86: 'SHL64', 0x87: 'SHRS64', 0x88: 'SHRU64', 0x51: 'EQ64', 0x52: 'NE64', 0x53: 'LTS64', 0x55: 'GTS64',
         0x57: 'LES64', 0x59: 'GES64', 0x50: 'EQZ64', 0xA7: 'WRAP', 0xAC: 'EXTS', 0xAD: 'EXTU'}

class Decoder:
    def __init__(self, m):
        self.m = m
        self.code = []     # [name, A, B]
        self.kj = []       # i64 constants
        self.bp = []       # br_table pool: pc, sp-offset pairs
        self.fe = []; self.fnp = []; self.fnl = []
        self.fix = []      # (code index, field, func k, wasm instr index, 'at'|'after')

    def func(self, k):
        m = self.m
        ltypes, results, ann = wasm.analyze(m, k)
        ins = m.codes[k][1]
        nl = len(ltypes)
        self.fe.append(len(self.code)); self.fnp.append(len(m.types[m.func_types[k]][0])); self.fnl.append(nl)
        at = {}; after = {}
        pend = []
        def emit(name, A=0, B=0):
            self.code.append([name, A, B]); return len(self.code) - 1
        def tpos(tgt, a):
            # wasm index whose 'at' pc is the branch destination
            if tgt['kind'] == 'loop':
                return ('at', tgt['start'])
            return ('at', ann[tgt['start']]['end'])
        def br(tgt, a, conditional):
            hcur = a['h'] - (1 if conditional else 0)
            if tgt['kind'] == 'func':
                assert not conditional, 'conditional return'
                emit('RET1' if results else 'RET0'); return
            ar = len(tgt['arity'])
            adj = hcur != tgt['height'] + ar
            name = ('BRIF' if conditional else 'BR') + (('ADJ%d' % ar) if adj else '')
            c = emit(name, 0, nl + tgt['height'])
            pend.append((c, 1) + tpos(tgt, a))
        for i, x in enumerate(ins):
            op = x[0]; a = ann[i]
            at[i] = len(self.code)
            if a['dead']:
                after[i] = len(self.code); continue
            if op in (BLOCK, LOOP, 0x01):
                pass
            elif op == IF:
                c = emit('BRZ')
                pend.append((c, 1, 'after', a['else']) if a['else'] is not None else (c, 1, 'at', a['end']))
            elif op == ELSE:
                c = emit('BR'); pend.append((c, 1, 'at', ann[a['start']]['end']))
            elif op == END:
                if a['start'] is None:
                    emit('RET1' if results else 'RET0')
            elif op == 0x0C:
                br(a['br'], a, False)
            elif op == 0x0D:
                tgt = a['br']
                if tgt['kind'] == 'func':
                    c = emit('BRZ'); emit('RET1' if results else 'RET0'); self.code[c][1] = len(self.code)
                else:
                    br(tgt, a, True)
            elif op == 0x0E:
                base = len(self.bp)
                for tgt in a['brs'] + [a['brd']]:
                    assert tgt['kind'] != 'func' and not tgt['arity']
                    self.bp += [0, nl + tgt['height']]
                    pend.append((len(self.bp) - 2, 'bp') + tpos(tgt, a))
                emit('BRT', base, len(a['brs']))
            elif op == 0x0F:
                emit('RET1' if results else 'RET0')
            elif op == 0x00:
                emit('UNREACH')
            elif op == 0x10:
                assert x[1] >= m.nimports
                emit('CALL', x[1] - m.nimports)
            elif op == 0x11:
                emit('CALLI', x[1])
            elif op == 0x1A:
                emit('DROP')
            elif op == 0x1B:
                emit('SELECT64' if a['ty'] == I64 else 'SELECT')
            elif op in (0x20, 0x21, 0x22):
                emit({0x20: 'LGET', 0x21: 'LSET', 0x22: 'LTEE'}[op] + ('64' if ltypes[x[1]] == I64 else ''), x[1])
            elif op == 0x23:
                emit('GGET', x[1])
            elif op == 0x24:
                emit('GSET', x[1])
            elif op == 0x28:
                emit('LOAD32' if x[1] >= 2 else 'LOAD32X', x[2])
            elif op == 0x36:
                emit('STORE32' if x[1] >= 2 else 'STORE32X', x[2])
            elif op in (0x29, 0x2C, 0x2D, 0x2E, 0x2F, 0x37, 0x3A, 0x3B):
                emit({0x29: 'LOAD64', 0x2C: 'LOAD8S', 0x2D: 'LOAD8U', 0x2E: 'LOAD16S', 0x2F: 'LOAD16U',
                      0x37: 'STORE64', 0x3A: 'STORE8', 0x3B: 'STORE16'}[op], x[2])
            elif op == 0x41:
                emit('CONST', w2vba.s32(x[1]))
            elif op == 0x42:
                self.kj.append(x[1]); emit('CONST64', len(self.kj) - 1)
            elif op in BIN32:
                emit(BIN32[op])
            elif op in BIN64:
                emit(BIN64[op])
            else:
                raise w2vba.Unsupported(wasm.opclass(op))
            after[i] = len(self.code)
        for c, field, how, idx in pend:
            pc = at[idx] if how == 'at' else after[idx]
            if field == 'bp':
                self.bp[c] = pc
            else:
                self.code[c][field] = pc

def build(path):
    m = wasm.parse(path)
    D = Decoder(m)
    for k in range(len(m.codes)):
        D.func(k)
    T = w2vba.Translator(path, 'L')
    datasubs, datanames = T.data_init()
    L = ['Option Explicit', w2vba.RUNTIME_DECL, w2vba.RUNTIME_L_DECL]
    n = len(D.code)
    L.append('Public OP(0 To %d) As Long, IA(0 To %d) As Long, IB(0 To %d) As Long' % (n, n, n))
    L.append('Public S(0 To 65535) As Long, SJ(0 To 65535) As LongLong, CS(0 To 8191) As Long')
    L.append('Public G(0 To %d) As Long, KJ(0 To %d) As LongLong, BP(0 To %d) As Long' % (max(len(m.globals), 1), max(len(D.kj), 1), max(len(D.bp), 1)))
    nf = len(m.codes)
    L.append('Public FE(0 To %d) As Long, FNP(0 To %d) As Long, FNL(0 To %d) As Long, TBL(0 To %d) As Long' % (nf, nf, nf, max(m.table_size, 1)))
    L += [w2vba.RUNTIME, w2vba.RUNTIME_L]
    stm = []
    for pc, (name, A, B) in enumerate(D.code):
        s = 'OP(%d) = %d' % (pc, OPNUM[name])
        if A: s += ': IA(%d) = %s' % (pc, w2vba.lit32(A))
        if B: s += ': IB(%d) = %s' % (pc, w2vba.lit32(B))
        stm.append(s)
    for i, v in enumerate(D.kj):
        stm.append('KJ(%d) = %s' % (i, w2vba.lit64(v)))
    for i, v in enumerate(D.bp):
        stm.append('BP(%d) = %d' % (i, v))
    for k in range(nf):
        stm.append('FE(%d) = %d: FNP(%d) = %d: FNL(%d) = %d' % (k, D.fe[k], k, D.fnp[k], k, D.fnl[k]))
    for off, fs in m.elems:
        for j, f in enumerate(fs):
            stm.append('TBL(%d) = %d' % (off + j, f - m.nimports))
    for gi, (t, mut, v) in enumerate(m.globals):
        stm.append('G(%d) = %s' % (gi, w2vba.lit32(v)))
    names = []
    for i in range(0, len(stm), 300):
        nm = 'ProgInit%d' % (i // 300); names.append(nm)
        L += ['Private Sub %s()' % nm] + stm[i:i + 300] + ['End Sub']
    L += datasubs
    L += ['Public Sub RtInit()', 'RtTables', 'ReDim M(0 To %d)' % (m.mem[0] * 16384 - 1)] + names + datanames + ['End Sub']
    # the interpreter
    order = list(H)
    L += ['Public Function Exec(ByVal f As Long, ByVal nargs As Long, ByVal a0 As Long) As Long',
          'Dim pc As Long, sp As Long, fp As Long, csp As Long, a As Long, i As Long, k As Long, x As LongLong',
          'S(0) = a0: fp = 0: sp = nargs',
          'For k = sp To FNL(f) - 1: S(k) = 0: SJ(k) = 0: Next k',
          'sp = FNL(f): CS(0) = -1: CS(1) = 0: csp = 2: pc = FE(f)',
          'DSP:']
    labs = ['H%d' % OPNUM[nm] for nm in order]
    chunks = [', '.join(labs[i:i + 12]) for i in range(0, len(labs), 12)]
    L.append('On OP(pc) GoTo ' + ', _\n'.join(chunks))
    L.append('Trap')
    for nm in order:
        lines = H[nm]
        L.append('H%d:' % OPNUM[nm])
        L += lines
        if not lines[-1].endswith('GoTo DSP') or lines[-1].startswith('sp = sp - 1: If') or ' If ' in lines[-1] or lines[-1].startswith('If'):
            L.append('pc = pc + 1: GoTo DSP')
    L += ['DOCALL:',
          'CS(csp) = pc + 1: CS(csp + 1) = fp: csp = csp + 2: fp = sp - FNP(f)',
          'For k = sp To fp + FNL(f) - 1: S(k) = 0: SJ(k) = 0: Next k',
          'sp = fp + FNL(f): pc = FE(f): GoTo DSP',
          'DORET:',
          'csp = csp - 2: fp = CS(csp + 1): pc = CS(csp)',
          'If pc >= 0 Then GoTo DSP',
          'Exec = S(0)',
          'End Function']
    exp = {}
    for name, (kind, idx) in m.exports.items():
        if kind == 0:
            pt, rt = func_type(m, idx)
            k = idx - m.nimports
            if pt:
                L += ['Public Function x_%s(ByVal a0 As Long) As Long' % name, 'x_%s = Exec(%d, 1, a0)' % (name, k), 'End Function']
            else:
                L += ['Public Sub x_%s()' % name, 'Dim r As Long: r = Exec(%d, 0, 0)' % k, 'End Sub']
    return [('Rt', '\n'.join(L) + '\n')], dict(ops=len(D.code), handlers=len(H))
