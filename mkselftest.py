"""Self-test workbook: runs the hand-written runtime helpers and a few VBA-semantics
probes in real Excel, so the results can be compared with the offline emulator."""
import random, sys
import mkxlsm, rtvba, vbaemu, hostvba
from w2v import lit32, lit64

random.seed(11)
def r64():
    k = random.randrange(4)
    if k == 0: v = random.getrandbits(64)
    elif k == 1: v = random.choice([0, 1, 2**64 - 1, 2**63 - 1, 2**63, 2**32, 2**64 - 2**32, 2**31, 0xFFFFFFFF, 65536])
    elif k == 2: v = random.getrandbits(random.randrange(1, 64))
    else: v = 2**64 - random.getrandbits(random.randrange(1, 63)) - 1
    v &= 2**64 - 1
    return v - 2**64 if v >> 63 else v
def r32():
    v = r64() & 0xFFFFFFFF
    return v - 2**32 if v >> 31 else v

calls = []
for _ in range(40):
    a, b, k = r64(), r64(), random.randrange(64)
    calls += ['Add64(%s, %s)' % (lit64(a), lit64(b)), 'Sub64(%s, %s)' % (lit64(a), lit64(b)), 'Mul64(%s, %s)' % (lit64(a), lit64(b)),
              'Shl64(%s, %d)' % (lit64(a), k), 'ShrU64(%s, %d)' % (lit64(a), k), 'ShrS64(%s, %d)' % (lit64(a), k), 'Rotl64(%s, %d)' % (lit64(a), k),
              'Clz64(%s)' % lit64(a), 'Ctz64(%s)' % lit64(a), 'Popcnt64(%s)' % lit64(a)]
    if b: calls += ['DivU64(%s, %s)' % (lit64(a), lit64(b)), 'RemU64(%s, %s)' % (lit64(a), lit64(b))]
    x, y = r32(), r32()
    calls += ['W32(%s)' % lit64(a), 'ShrS32(%s, %d)' % (lit32(x), k), 'ShrU32(%s, %d)' % (lit32(x), k), 'Rotl32(%s, %d)' % (lit32(x), k),
              'Rotr32(%s, %d)' % (lit32(x), k), 'Clz32(%s)' % lit32(x), 'Ctz32(%s)' % lit32(x), 'Popcnt32(%s)' % lit32(x)]
    if y: calls += ['DivU32(%s, %s)' % (lit32(x), lit32(y)), 'RemU32(%s, %s)' % (lit32(x), lit32(y)),
                    '(%s \\ %s)' % (lit32(x), lit32(y)), '(%s Mod %s)' % (lit32(x), lit32(y))]
    calls += ['CLng(((CLngLng(%s) * %s And 4294967295^) Xor 2147483648^) - 2147483648^)' % (lit32(x), lit32(y)),
              '((%s And %s) \\ %d&)' % (lit32(x), lit32(~((1 << (k & 31)) - 1)), 1 << (k & 31)) if (k & 31) < 31 else 'CLng(%s < 0)' % lit32(x),
              '((%s Xor &H80000000) < (%s Xor &H80000000)) And 1&' % (lit32(x), lit32(y)),
              '((%s Xor P2(63)) < (%s Xor P2(63))) And 1&' % (lit64(a), lit64(b)),
              '((%s And %s) \\ %d^)' % (lit64(a), lit64(~((1 << (k % 63)) - 1)), 1 << (k % 63))]

MEMPROG = '''    ST32 100, &H12345678: ST8 101, 255: ST16 103, -2: ST64 110, -81985529216486896^
    MemFill 200, 171, 37: MemCopy 205, 100, 20: MemCopy 101, 100, 30
    s = s & LD32(100) & "," & LD8(103) & "," & LD16(103) & "," & LD64(110) & "," & LD32(201) & "," & LD32(206) & "," & LD64(105) & "," & LD32(233) & ","
'''
chunks = [calls[i:i + 60] for i in range(0, len(calls), 60)]
src = ['Option Explicit', 'Public MemPages As Long']
for ci, ch in enumerate(chunks):
    src.append('Public Function T%d() As String' % ci)
    src.append('    Dim s As String')
    for c in ch:
        src.append('    s = s & (%s) & ","' % c)
    src.append('    T%d = s' % ci)
    src.append('End Function')
src.append('''Public Function SelfTest() As String
    Dim s As String
    On Error GoTo Fail
    RtInit 4
    MemPages = 2
%s    s = s & F64Bits(1.5) & "," & F32Bits(-2.25) & "," & BitsF64(4611686018427387904^) & "," & BitsF32(1065353216) & ","
%s    SelfTest = s
    Exit Function
Fail:
    SelfTest = s & "ERROR " & Err.Number & " " & Err.Description
End Function
''' % (MEMPROG, ''.join('    s = s & T%d()\n' % i for i in range(len(chunks)))))
SELF = '\n'.join(src) + '\n'

def expected():
    # same code through the emulator: strings are not in the subset, so evaluate each call directly
    body = ['Option Explicit', 'Public MemPages As Long', 'Public R(0 To %d) As LongLong' % (len(calls) + 20), 'Public Sub Go()']
    for i, c in enumerate(calls):
        body.append('R(%d) = %s' % (i, c))
    body.append('End Sub')
    mem = '''Public Sub Mem()
ST32 100, &H12345678: ST8 101, 255: ST16 103, -2: ST64 110, -81985529216486896^
MemFill 200, 171, 37: MemCopy 205, 100, 20: MemCopy 101, 100, 30
R(0) = LD32(100): R(1) = LD8(103): R(2) = LD16(103): R(3) = LD64(110): R(4) = LD32(201): R(5) = LD32(206): R(6) = LD64(105): R(7) = LD32(233)
End Sub
'''
    e = vbaemu.Emu([('T', '\n'.join(body) + '\n' + mem), ('Rt', rtvba.RT)])
    e.call('RtInit', 4); e.ns['MemPages'] = 2
    e.call('Mem'); memvals = [str(v) for v in e.ns['R'][:8]]
    e.call('Go')
    vals = [str(v) for v in e.ns['R'][:len(calls)]]
    return memvals + ['4609434218613702656', '-1072693248', '2', '1'] + vals

if __name__ == '__main__':
    mkxlsm.build_xlsm('out/selftest.xlsm', [('Rt', rtvba.RT), ('RtF', rtvba.RTF), ('SelfTest', SELF)])
    exp = expected()
    import xl
    out, dt = xl.session('out/selftest.xlsm', 'SelfTest.SelfTest', timeout=120)
    got = out.rstrip(',').split(',')
    print('excel returned %d values in %.1fs; expected %d' % (len(got), dt, len(exp)))
    if 'ERROR' in out or out.startswith(('DIALOG', 'TIMEOUT')):
        print(out[-300:])
    bad = [(i, g, e) for i, (g, e) in enumerate(zip(got, exp)) if g != e]
    print('mismatches:', len(bad) + abs(len(got) - len(exp)))
    for i, g, e in bad[:10]:
        what = calls[i - 12] if i >= 12 else 'mem/float #%d' % i
        print('   ', what, '| excel', g, '| emulator', e)
