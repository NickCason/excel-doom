"""Pre-flight check for generated VBA: every identifier used in a procedure must be
declared there or at module level somewhere in the project (case-insensitively).
Catches 'Variable not defined' before Excel puts up a dialog."""
import re, sys

KEYWORDS = set('''option explicit public private function sub end dim as long longlong byte double string boolean
if then else elseif goto select case for to next step exit and or xor not mod byval byref on error err raise
redim const call true false clng clnglng cdbl format timer vbobjecterror number description do loop while wend
integer single variant is each in with set let resume'''.split())

def lint(mods):
    glob = set()
    for _, src in mods:
        for line in src.split('\n'):
            m = re.match(r'\s*(?:Public|Private)\s+(?:Function|Sub)\s+(\w+)', line)
            if m:
                glob.add(m.group(1).lower()); continue
            m = re.match(r'\s*(?:Public|Private)\s+(.*)', line)
            if m and not re.match(r'(Function|Sub)\b', m.group(1)):
                for d in m.group(1).split(','):
                    nm = re.match(r'\s*(\w+)', d)
                    if nm: glob.add(nm.group(1).lower())
    problems = []
    for mname, src in mods:
        seen_proc = False
        for ln, line in enumerate(src.split('\n'), 1):
            if re.match(r'\s*(?:Public|Private)\s+(Function|Sub)\b', line):
                seen_proc = True
            elif seen_proc and re.match(r'\s*(?:Public|Private|Dim|Const)\s+\w+.*\bAs\b', line) and not line.startswith(' '):
                if re.match(r'\s*(?:Public|Private)\s', line):
                    problems.append((mname, None, ln, 'DECL-AFTER-PROC', line.strip()[:110]))
    for mname, src in mods:
        local = None; pname = None
        for ln, line in enumerate(src.replace(' _\n', ' ').split('\n'), 1):
            m = re.match(r'\s*(?:Public|Private)\s+(Function|Sub)\s+(\w+)\((.*?)\)', line)
            if m:
                pname = m.group(2); local = {pname.lower()}
                for p in m.group(3).split(','):
                    nm = re.match(r'\s*(?:ByVal|ByRef)?\s*(\w+)', p)
                    if nm: local.add(nm.group(1).lower())
                continue
            if local is None:
                continue
            if re.match(r'\s*End (Function|Sub)', line):
                local = None; continue
            code = re.sub(r'"[^"]*"', '""', line)
            code = re.sub(r"'.*$", '', code)
            m = re.match(r'\s*Dim\s+(.*)', code)
            if m:
                for d in m.group(1).split(','):
                    nm = re.match(r'\s*(\w+)', d)
                    if nm: local.add(nm.group(1).lower())
                m2 = re.search(r':(.*)$', code)
                if not m2: continue
                code = m2.group(1)
            for lab in re.findall(r'^\s*(\w+):\s*$', code):
                local.add(lab.lower())
            code = re.sub(r'&H[0-9A-Fa-f]+&?', ' ', code)
            code = re.sub(r'\b\d+[&^#]?', ' ', code)
            code = re.sub(r'GoTo\s+\w+(\s*,\s*\w+)*', ' ', code)
            for tok in re.findall(r'[A-Za-z_]\w*', code):
                t = tok.lower()
                if t in KEYWORDS or t in local or t in glob:
                    continue
                if re.fullmatch(r'[lex]\d+|h\d+', t) and code.strip().endswith(':'):
                    continue
                problems.append((mname, pname, ln, tok, line.strip()[:110]))
    return problems

if __name__ == '__main__':
    pass
