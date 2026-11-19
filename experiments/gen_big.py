"""Synthetic load test: VBA shaped like translator output, at a chosen total line count."""
import sys, random
sys.path.insert(0, '.')
import mkxlsm

def gen(total_lines, n_modules, lines_per_fn, out):
    random.seed(7)
    fns_per_mod = max(1, total_lines // n_modules // lines_per_fn)
    mods = []
    count = 0
    for m in range(n_modules):
        L = ['Option Explicit', '']
        for f in range(fns_per_mod):
            name = 'f%d_%d' % (m, f)
            L.append('Public Function %s(ByVal p0 As Long, ByVal p1 As Long) As Long' % name)
            L.append('Dim s0 As Long, s1 As Long, s2 As Long, s3 As Long, l2 As Long, l3 As Long')
            L.append('Dim w As LongLong')
            L.append('s0 = p0: s1 = p1')
            body = lines_per_fn - 7
            lab = 0
            for i in range(body):
                k = random.randrange(8)
                a, b, c = (random.choice(['s0', 's1', 's2', 's3', 'l2', 'l3']) for _ in range(3))
                if k == 0:
                    L.append('w = CLngLng(%s) + %s: %s = CLng(((w + 2147483648^) And 4294967295^) - 2147483648^)' % (a, b, c))
                elif k == 1:
                    L.append('%s = %s Xor %d' % (c, a, random.randrange(1 << 30)))
                elif k == 2:
                    L.append('%s = Mem((%s And &HFFFC&) \\ 4)' % (c, a))
                elif k == 3:
                    L.append('Mem((%s And &HFFFC&) \\ 4) = %s' % (a, b))
                elif k == 4:
                    L.append('If %s < %s Then GoTo L%d' % (a, b, lab))
                    L.append('%s = %s And %d' % (c, a, random.randrange(1 << 30)))
                    L.append('L%d:' % lab)
                    lab += 1
                elif k == 5:
                    L.append('%s = (%s And &H7FFFFFFF) \\ %d' % (c, a, 1 << random.randrange(1, 20)))
                elif k == 6:
                    L.append('w = CLngLng(%s) * %d: %s = CLng(((w + 2147483648^) And 4294967295^) - 2147483648^)' % (a, random.randrange(3, 99), c))
                else:
                    L.append('%s = %s Or (%s And 255)' % (c, a, b))
            L.append('%s = s0 Xor s1 Xor s2 Xor s3' % name)
            L.append('End Function')
            L.append('')
        L.append('Public Function CallAll%d(ByVal x As Long) As Long' % m)
        for f in range(fns_per_mod):
            L.append('x = x Xor f%d_%d(%d, x)' % (m, f, f + 1))
        L.append('CallAll%d = x' % m)
        L.append('End Function')
        count += len(L)
        mods.append(('T%02d' % m, '\n'.join(L) + '\n'))
    drv = ['Option Explicit', 'Public Mem(0 To 16383) As Long', '',
           'Public Function RunAll() As String',
           'Dim t As Double, x As Long, i As Long', 't = Timer']
    for m in range(n_modules):
        drv.append('x = CallAll%d(x)' % m)
    drv.append('Dim t1 As Double: t1 = Timer - t')
    drv.append('t = Timer')
    drv.append('For i = 1 To 20')
    for m in range(n_modules):
        drv.append('x = CallAll%d(x)' % m)
    drv.append('Next i')
    drv.append('RunAll = "firstcall(compile)=" & Format(t1, "0.00") & "s warm20=" & Format(Timer - t, "0.000") & "s x=" & x')
    drv.append('End Function')
    count += len(drv)
    mods.append(('Driver', '\n'.join(drv) + '\n'))
    vba = mkxlsm.build_xlsm(out, mods)
    print('%s: %d lines, %d modules x %d fns x %d lines, vbaProject %.1f MB' % (out, count, n_modules, fns_per_mod, lines_per_fn, len(vba) / 1e6))

if __name__ == '__main__':
    gen(int(sys.argv[1]), int(sys.argv[2]), int(sys.argv[3]), sys.argv[4])
