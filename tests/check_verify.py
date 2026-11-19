"""Compare the string returned by the workbook's Host.Verify macro with the expected
frame checksums (tests/verify_sums.txt). Usage: python tests/check_verify.py result.txt"""
import sys, os
text = open(sys.argv[1], encoding='utf-8', errors='replace').read()
head, _, sums = text.partition('|sums=')
got = [int(x) for x in sums.strip().rstrip(',').split(',') if x.strip()]
want = [int(x) for x in open(os.path.join(os.path.dirname(__file__), 'verify_sums.txt')).read().split()]
same = 0
for a, b in zip(got, want):
    if a != b:
        break
    same += 1
print('macro said:', head[:200])
print('frames returned: %d, expected: %d, identical from the start: %d' % (len(got), len(want), same))
sys.exit(0 if same == len(want) == len(got) else 1)
