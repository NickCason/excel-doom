"""Runs the small VBA subset our generators emit, by compiling it to Python.
Models VBA integer typing (Integer/Long/LongLong overflow raises), truncating
division, True = -1, and GoTo. Lets the generated VBA be tested without Excel."""
import re, math, struct, hashlib, marshal, os, sys

class VBAError(Exception):
    pass

RANK = {'B': 0, 'I': 1, 'L': 2, 'Q': 3, 'S': 4, 'D': 5}
TYPES = {'Integer': 'I', 'Long': 'L', 'LongLong': 'Q', 'Single': 'S', 'Double': 'D', 'Boolean': 'B'}
BITS = {'I': 16, 'L': 32, 'Q': 64, 'B': 16}

TOK = re.compile(r'''\s*(?:
  (?P<hex>&H[0-9A-Fa-f]+[&^]?) |
  (?P<num>\d+\.\d*(?:[eE][+-]?\d+)?[\#!]? | \d+[eE][+-]?\d+[\#!]? | \d+[&^\#!%]?) |
  (?P<str>"(?:[^"]|"")*") |
  (?P<id>[A-Za-z_][A-Za-z_0-9]*) |
  (?P<op><>|<=|>=|[-+*\\/=<>(),.])
)''', re.X)

def tokenize(s):
    out = []; i = 0; s = s.rstrip()
    while i < len(s):
        m = TOK.match(s, i)
        if not m or m.end() == i:
            raise VBAError('cannot tokenize: %r at %d' % (s, i))
        out.append((m.lastgroup, m.group(m.lastgroup)))
        i = m.end()
    return out

def maxt(a, b):
    t = a if RANK[a] >= RANK[b] else b
    return 'I' if t == 'B' else t

class Sym:
    def __init__(self):
        self.vars = {}      # module-level: name -> (type, is_array)
        self.funcs = {}     # name -> ([param types], ret type or None)

class Expr:
    """Typed expression compiler over a token list."""
    def __init__(self, toks, fn):
        self.t = toks; self.i = 0; self.fn = fn

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def eat(self, val=None):
        k, v = self.peek()
        if val is not None and (v is None or v.lower() != val.lower()):
            raise VBAError('expected %r got %r in %r' % (val, v, ' '.join(x[1] for x in self.t)))
        self.i += 1
        return k, v

    def kw(self, word):
        k, v = self.peek()
        return k == 'id' and v.lower() == word

    # precedence climbing: Xor < Or < And < Not < cmp < +- < Mod < \ < */ < unary -
    def parse(self):
        return self.p_xor()

    def p_xor(self):
        a = self.p_or()
        while self.kw('xor'):
            self.eat(); b = self.p_or(); a = self.bit('^', a, b)
        return a

    def p_or(self):
        a = self.p_and()
        while self.kw('or'):
            self.eat(); b = self.p_and(); a = self.bit('|', a, b)
        return a

    def p_and(self):
        a = self.p_not()
        while self.kw('and'):
            self.eat(); b = self.p_not(); a = self.bit('&', a, b)
        return a

    def p_not(self):
        if self.kw('not'):
            self.eat(); a = self.p_not()
            if a[2] is not None:
                return ('(-1 if not (%s) else 0)' % a[2], 'B', 'not (%s)' % a[2])
            return ('(~%s)' % a[0], a[1], None)
        return self.p_cmp()

    def p_cmp(self):
        a = self.p_add()
        while self.peek()[1] in ('=', '<>', '<', '>', '<=', '>='):
            op = self.eat()[1]; b = self.p_add()
            py = {'=': '==', '<>': '!='}.get(op, op)
            c = '%s %s %s' % (a[0], py, b[0])
            a = ('(-1 if %s else 0)' % c, 'B', c)
        return a

    def p_add(self):
        a = self.p_mod()
        while self.peek()[1] in ('+', '-'):
            op = self.eat()[1]; b = self.p_mod(); a = self.arith(op, a, b)
        return a

    def p_mod(self):
        a = self.p_idiv()
        while self.kw('mod'):
            self.eat(); b = self.p_idiv()
            t = maxt(a[1], b[1])
            a = ('_imod(%s, %s)' % (a[0], b[0]), t, None)
        return a

    def p_idiv(self):
        a = self.p_mul()
        while self.peek()[1] == '\\':
            self.eat(); b = self.p_mul()
            t = maxt(a[1], b[1])
            a = (self.chk('_idiv(%s, %s)' % (a[0], b[0]), t), t, None)
        return a

    def p_mul(self):
        a = self.p_unary()
        while self.peek()[1] in ('*', '/'):
            op = self.eat()[1]; b = self.p_unary()
            if op == '/':
                a = ('(%s / %s)' % (a[0], b[0]), 'D', None)
            else:
                a = self.arith(op, a, b)
        return a

    def p_unary(self):
        if self.peek()[1] == '-':
            self.eat(); a = self.p_unary()
            if a[1] in 'SD':
                return ('(-%s)' % a[0], a[1], None)
            t = maxt(a[1], 'I')
            return (self.chk('(-%s)' % a[0], t), t, None)
        return self.p_primary()

    def chk(self, code, t):
        if self.fn.strict and t in BITS:
            return '_c%d(%s)' % (BITS[t], code)
        return code

    def arith(self, op, a, b):
        t = maxt(a[1], b[1])
        code = '(%s %s %s)' % (a[0], op, b[0])
        if t in 'SD':
            return (code, t, None)
        return (self.chk(code, t), t, None)

    def bit(self, op, a, b):
        t = a[1] if RANK[a[1]] >= RANK[b[1]] else b[1]
        return ('(%s %s %s)' % (a[0], op, b[0]), t, None)

    def p_primary(self):
        k, v = self.eat()
        if k == 'hex':
            suf = v[-1] if v[-1] in '&^' else ''
            digits = v[2:-1] if suf else v[2:]
            n = int(digits, 16)
            if suf == '^':
                return (str(n if n < 1 << 63 else n - (1 << 64)), 'Q', None)
            if suf == '&' or len(digits) > 4:
                return ('(%d)' % (n if n < 1 << 31 else n - (1 << 32)), 'L', None)
            return ('(%d)' % (n if n < 1 << 15 else n - (1 << 16)), 'I', None)
        if k == 'num':
            suf = v[-1] if v[-1] in '&^#!%' else ''
            body = v[:-1] if suf else v
            if suf in ('#', '!') or '.' in body or 'e' in body.lower():
                return (repr(float(body)), 'D', None)
            n = int(body)
            t = {'&': 'L', '^': 'Q', '%': 'I'}.get(suf) or ('I' if n <= 32767 else 'L' if n <= 2147483647 else 'D')
            if t == 'D':
                return (repr(float(n)), 'D', None)
            return (str(n), t, None)
        if v == '(':
            a = self.parse(); self.eat(')')
            return ('(%s)' % a[0], a[1], a[2])
        if k == 'id':
            return self.ident(v)
        raise VBAError('unexpected token %r in %r' % (v, ' '.join(x[1] for x in self.t)))

    def args(self):
        out = []
        self.eat('(')
        if self.peek()[1] != ')':
            out.append(self.parse())
            while self.peek()[1] == ',':
                self.eat(); out.append(self.parse())
        self.eat(')')
        return out

    def ident(self, name):
        fn = self.fn; lo = name.lower()
        if lo == 'true': return ('(-1)', 'B', None)
        if lo == 'false': return ('0', 'B', None)
        call = self.peek()[1] == '('
        if call and lo in BUILTINS:
            a = self.args()
            return BUILTINS[lo](self, a)
        if name in fn.locals:
            t, arr = fn.locals[name]
            if arr:
                a = self.args(); return (fn.index(name, a[0][0]), t, None)
            return (name, t, None)
        if name == fn.name and not call:
            return ('_r', fn.ret, None)
        sym = fn.sym
        if name in sym.vars:
            t, arr = sym.vars[name]
            if arr:
                a = self.args(); return (fn.index(name, a[0][0]), t, None)
            return (name, t, None)
        if name in sym.funcs:
            pts, rt = sym.funcs[name]
            a = self.args() if call else []
            if len(a) != len(pts):
                raise VBAError('%s: %d args, expects %d (in %s)' % (name, len(a), len(pts), fn.name))
            if rt is None:
                raise VBAError('%s used as function but is a Sub (in %s)' % (name, fn.name))
            return ('%s(%s)' % (name, ', '.join(conv(x, pt, fn.strict) for x, pt in zip(a, pts))), rt, None)
        raise VBAError('undefined name %r in %s' % (name, fn.name))

def conv(e, to, strict=True):
    code, t = e[0], e[1]
    if to in ('L', 'Q', 'I'):
        if t in 'SD':
            return '_c%d(round(%s))' % (BITS[to], code)
        if RANK[t] > RANK[to] and strict:
            return '_c%d(%s)' % (BITS[to], code)
        return code
    if to == 'D':
        return code if t in 'SD' else 'float(%s)' % code
    if to == 'S':
        return '_f32(%s)' % code
    if to == 'B':
        return code if t == 'B' else '(-1 if %s else 0)' % code
    raise VBAError('conv to ' + to)

def _b1(name, fmt, rt):
    def f(ex, a):
        return (fmt % a[0][0], rt if rt else a[0][1], None)
    return f

BUILTINS = {
    'clng': lambda ex, a: (conv(a[0], 'L', True), 'L', None),
    'clnglng': lambda ex, a: (conv(a[0], 'Q', True), 'Q', None),
    'cdbl': lambda ex, a: (conv(a[0], 'D'), 'D', None),
    'csng': lambda ex, a: (conv(a[0], 'S'), 'S', None),
    'abs': _b1('abs', 'abs(%s)', None),
    'sqr': _b1('sqr', 'math.sqrt(%s)', 'D'),
    'int': _b1('int', 'float(math.floor(%s))', 'D'),
    'fix': _b1('fix', 'float(math.trunc(%s))', 'D'),
}

class FnGen:
    def __init__(self, emu, name, params, ret, body, strict):
        self.emu = emu; self.sym = emu.sym; self.name = name; self.ret = ret; self.strict = strict
        self.locals = {p: (t, False) for p, t in params}
        self.params = params; self.body = body

    def index(self, name, idx):
        return '%s[_ix(%s)]' % (name, idx) if self.emu.check_index else '%s[%s]' % (name, idx)

    def expr(self, text):
        ex = Expr(tokenize(text), self)
        r = ex.parse()
        if ex.i != len(ex.t):
            raise VBAError('trailing tokens in %r (%s)' % (text, self.name))
        return r

    def cond(self, text):
        r = self.expr(text)
        return r[2] if r[2] is not None else '(%s) != 0' % r[0]

    def generate(self):
        stmts = []      # (kind, payload)
        for line in self.body:
            line = line.strip()
            if not line or line.startswith("'"):
                continue
            if re.fullmatch(r'[A-Za-z_]\w*:', line):
                stmts.append(('label', line[:-1])); continue
            m = re.match(r'If (.+?) Then (.+)$', line)
            if m:
                rest = m.group(2)
                parts = re.split(r' Else ', rest, maxsplit=1)
                stmts.append(('slif', (m.group(1), [s.strip() for s in parts[0].split(':')],
                                       [s.strip() for s in parts[1].split(':')] if len(parts) > 1 else None)))
                continue
            for s in line.split(':'):
                s = s.strip()
                if s:
                    stmts.append(('s', s))
        labels = [p for k, p in stmts if k == 'label']
        self.labelno = {l: i + 1 for i, l in enumerate(labels)}
        self.has_goto = bool(labels)
        out = []; self.out = out
        self.ind = 1
        self.globals_assigned = set()
        self.loop_depth = 0
        segs = [[]]
        for k, p in stmts:
            if k == 'label':
                segs.append([])
            else:
                segs[-1].append((k, p))
        if self.has_goto:
            self.w('_pc = 0'); self.w('while True:'); self.ind += 1
            for si, seg in enumerate(segs):
                self.w('if _pc <= %d:' % si); self.ind += 1
                n0 = len(out)
                self.block(seg)
                if len(out) == n0:
                    self.w('pass')
                self.ind -= 1
            self.w('break'); self.ind -= 1
        else:
            self.block(segs[0])
        head = ['def %s(%s):' % (self.name, ', '.join(p for p, _ in self.params))]
        if self.globals_assigned:
            head.append('    global ' + ', '.join(sorted(self.globals_assigned)))
        init = []
        for n, (t, arr) in self.locals.items():
            if n in dict(self.params):
                continue
            init.append('    %s = %s' % (n, '[]' if arr else ('0.0' if t in 'SD' else '0')))
        init.append('    _r = %s' % ('0.0' if self.ret in ('S', 'D') else '0'))
        return '\n'.join(head + init + out + ['    return _r' if self.ret else '    return None']) + '\n'

    def w(self, s):
        self.out.append('    ' * self.ind + s)

    def block(self, stmts):
        """Structured statements; stack tracks open VBA blocks."""
        stack = []
        for k, p in stmts:
            if k == 'slif':
                c, th, el = p
                self.w('if %s:' % self.cond(c)); self.ind += 1
                for s in th: self.simple(s)
                self.ind -= 1
                if el is not None:
                    self.w('else:'); self.ind += 1
                    for s in el: self.simple(s)
                    self.ind -= 1
                continue
            s = p
            lo = s.lower()
            if lo.startswith('if ') and lo.endswith(' then'):
                self.w('if %s:' % self.cond(s[3:-5])); self.ind += 1; stack.append('if')
            elif lo.startswith('elseif ') and lo.endswith(' then'):
                self.ind -= 1; self.w('elif %s:' % self.cond(s[7:-5])); self.ind += 1
            elif lo == 'else':
                self.ind -= 1; self.w('else:'); self.ind += 1; self.w('pass')
            elif lo == 'end if':
                self.ind -= 1; assert stack.pop() == 'if', self.name
            elif lo.startswith('select case '):
                e = self.expr(s[12:])
                self.w('_sel = %s' % e[0]); stack.append(['select', 0])
            elif lo.startswith('case '):
                top = stack[-1]
                if top[1]:
                    self.ind -= 1
                if lo == 'case else':
                    self.w('else:' if top[1] else 'if True:')
                else:
                    vals = [self.expr(v.strip())[0] for v in s[5:].split(',')]
                    self.w('%s _sel in (%s,):' % ('elif' if top[1] else 'if', ', '.join(vals)))
                top[1] += 1
                self.ind += 1; self.w('pass')
            elif lo == 'end select':
                top = stack.pop(); assert top[0] == 'select'
                if top[1]:
                    self.ind -= 1
            elif lo.startswith('for '):
                m = re.match(r'For (\w+) = (.+?) To (.+?)(?: Step (.+))?$', s)
                v = m.group(1); a = self.expr(m.group(2))[0]; b = self.expr(m.group(3))[0]
                st = self.expr(m.group(4))[0] if m.group(4) else '1'
                self.note_assign(v)
                self.w('for %s in _range(%s, %s, %s):' % (v, a, b, st)); self.ind += 1; stack.append('for'); self.loop_depth += 1
            elif lo.startswith('next'):
                self.ind -= 1; assert stack.pop() == 'for'; self.loop_depth -= 1
            elif lo.startswith('do while '):
                self.w('while %s:' % self.cond(s[9:])); self.ind += 1; stack.append('do'); self.loop_depth += 1
            elif lo == 'loop':
                self.ind -= 1; assert stack.pop() == 'do'; self.loop_depth -= 1
            else:
                self.simple(s)
        assert not stack, (self.name, stack)

    def note_assign(self, name):
        if name not in self.locals and name in self.sym.vars:
            self.globals_assigned.add(name)

    def simple(self, s):
        lo = s.lower()
        if lo.startswith('dim '):
            for d in s[4:].split(','):
                m = re.match(r'\s*(\w+)(\(\))?\s+As\s+(\w+)', d)
                self.locals[m.group(1)] = (TYPES[m.group(3)], bool(m.group(2)))
            return
        if lo.startswith('redim '):
            m = re.match(r'ReDim (\w+)\((.+?) To (.+)\)$', s)
            name = m.group(1); hi = self.expr(m.group(3))[0]
            t = (self.locals.get(name) or self.sym.vars[name])[0]
            self.note_assign(name)
            self.w('%s = [%s] * (%s + 1)' % (name, '0.0' if t in 'SD' else '0', hi)); return
        if lo.startswith('goto '):
            assert self.loop_depth == 0, 'GoTo inside loop in ' + self.name
            self.w('_pc = %d' % self.labelno[s[5:].strip()]); self.w('continue'); return
        if lo in ('exit function', 'exit sub'):
            self.w('return _r' if self.ret else 'return None'); return
        if lo in ('exit do', 'exit for'):
            self.w('break'); return
        if lo.startswith('err.raise'):
            self.w('raise VBAError(%r)' % s); return
        toks = tokenize(s)
        # assignment?
        depth = 0; eq = None
        for i, (k, v) in enumerate(toks):
            if v == '(': depth += 1
            elif v == ')': depth -= 1
            elif v == '=' and depth == 0:
                eq = i; break
        name = toks[0][1]
        if eq is not None and (eq == 1 or (toks[1][1] == '(' and toks[eq - 1][1] == ')')) and name not in self.sym.funcs or (eq == 1 and name == self.name):
            rhs = Expr(toks[eq + 1:], self); r = rhs.parse()
            assert rhs.i == len(rhs.t), s
            if eq == 1:
                if name == self.name:
                    self.w('_r = %s' % conv(r, self.ret, self.strict)); return
                if name in self.locals:
                    t = self.locals[name][0]
                elif name in self.sym.vars:
                    t = self.sym.vars[name][0]; self.globals_assigned.add(name)
                else:
                    raise VBAError('assignment to undefined %r in %s' % (name, self.name))
                self.w('%s = %s' % (name, conv(r, t, self.strict))); return
            t = (self.locals.get(name) or self.sym.vars.get(name))
            if t is None:
                raise VBAError('undefined array %r in %s' % (name, self.name))
            ix = Expr(toks[2:eq - 1], self); ie = ix.parse()
            assert ix.i == len(ix.t), s
            self.w('%s = %s' % (self.index(name, ie[0]), conv(r, t[0], self.strict))); return
        # call statement
        if name not in self.sym.funcs:
            raise VBAError('unknown statement %r in %s' % (s, self.name))
        pts, rt = self.sym.funcs[name]
        args = []
        if len(toks) > 1:
            ex = Expr(toks[1:], self)
            args.append(ex.parse())
            while ex.peek()[1] == ',':
                ex.eat(); args.append(ex.parse())
            assert ex.i == len(ex.t), s
        if len(args) != len(pts):
            raise VBAError('%s: %d args, expects %d (in %s)' % (name, len(args), len(pts), self.name))
        self.w('%s(%s)' % (name, ', '.join(conv(x, pt, self.strict) for x, pt in zip(args, pts))))

def _mk_helpers(ns):
    def _c16(x):
        if -32768 <= x <= 32767: return x
        raise VBAError('Overflow (Integer): %r' % x)
    def _c32(x):
        if -2147483648 <= x <= 2147483647: return x
        raise VBAError('Overflow (Long): %r' % x)
    def _c64(x):
        if -9223372036854775808 <= x <= 9223372036854775807: return x
        raise VBAError('Overflow (LongLong): %r' % x)
    def _idiv(a, b):
        if b == 0: raise VBAError('Division by zero')
        q = abs(a) // abs(b)
        return q if (a < 0) == (b < 0) else -q
    def _imod(a, b):
        return a - b * _idiv(a, b)
    def _ix(i):
        if i < 0: raise VBAError('Subscript out of range: %r' % i)
        return i
    def _range(a, b, st):
        return range(a, b + (1 if st > 0 else -1), st)
    def _f32(x):
        return struct.unpack('<f', struct.pack('<f', x))[0]
    ns.update(_c16=_c16, _c32=_c32, _c64=_c64, _idiv=_idiv, _imod=_imod, _ix=_ix, _range=_range, _f32=_f32,
              VBAError=VBAError, math=math)
    ns['F64Bits'] = lambda x: struct.unpack('<q', struct.pack('<d', x))[0]
    ns['BitsF64'] = lambda x: struct.unpack('<d', struct.pack('<q', x))[0]
    ns['F32Bits'] = lambda x: struct.unpack('<i', struct.pack('<f', x))[0]
    ns['BitsF32'] = lambda x: struct.unpack('<f', struct.pack('<i', x))[0]

class Emu:
    def __init__(self, modules, natives=None, strict=True, check_index=True, cache_dir='work/emucache'):
        """modules: [(name, source)]. natives: {name: (pyfunc, [param types], ret type)}."""
        self.sym = Sym(); self.strict = strict; self.check_index = check_index
        ns = {}; self.ns = ns
        _mk_helpers(ns)
        self.sym.funcs.update({'F64Bits': (['D'], 'Q'), 'BitsF64': (['Q'], 'D'), 'F32Bits': (['S'], 'L'), 'BitsF32': (['L'], 'S')})
        for name, (f, pts, rt) in (natives or {}).items():
            ns[name] = f; self.sym.funcs[name] = (pts, rt)
        fns = []
        for mname, src in modules:
            lines = src.split('\n'); i = 0
            while i < len(lines):
                line = lines[i].strip(); i += 1
                if not line or line.startswith("'") or line.startswith('Option '):
                    continue
                m = re.match(r'(?:Public|Private) (Function|Sub) (\w+)\((.*?)\)(?: As (\w+))?$', line)
                if m:
                    params = []
                    for p in filter(None, (x.strip() for x in m.group(3).split(','))):
                        pm = re.match(r'ByVal (\w+) As (\w+)', p)
                        params.append((pm.group(1), TYPES[pm.group(2)]))
                    ret = TYPES[m.group(4)] if m.group(4) else None
                    body = []
                    while not re.match(r'End (Function|Sub)$', lines[i].strip()):
                        body.append(lines[i]); i += 1
                    i += 1
                    if m.group(2) in self.sym.funcs:
                        raise VBAError('duplicate procedure ' + m.group(2))
                    self.sym.funcs[m.group(2)] = ([t for _, t in params], ret)
                    fns.append((m.group(2), params, ret, body))
                    continue
                m = re.match(r'(?:Public|Private|Dim) (.+)$', line)
                if m:
                    for d in re.split(r',\s*(?![^()]*\))', m.group(1)):
                        dm = re.match(r'\s*(\w+)(?:\((?:(\d+) To (\d+))?\))?\s+As\s+(\w+)', d)
                        name, lo_, hi, t = dm.group(1), dm.group(2), dm.group(3), TYPES[dm.group(4)]
                        arr = '(' in d
                        if name in self.sym.vars:
                            raise VBAError('duplicate module variable ' + name)
                        self.sym.vars[name] = (t, arr)
                        zero = 0.0 if t in 'SD' else 0
                        ns[name] = ([zero] * (int(hi) + 1) if hi else []) if arr else zero
                    continue
                raise VBAError('unsupported module-level line in %s: %r' % (mname, line))
        low = {}
        for n in list(self.sym.vars) + list(self.sym.funcs):
            if n.lower() in low and low[n.lower()] != n:
                raise VBAError('names differ only by case: %s / %s' % (n, low[n.lower()]))
            low[n.lower()] = n
        srcs = [FnGen(self, name, params, ret, body, strict).generate() for name, params, ret, body in fns]
        full = '\n'.join(srcs)
        self.pysrc_len = len(full)
        key = hashlib.sha1((full + sys.version).encode()).hexdigest()
        os.makedirs(cache_dir, exist_ok=True)
        cpath = os.path.join(cache_dir, key + '.bin')
        if os.path.exists(cpath):
            code = marshal.load(open(cpath, 'rb'))
        else:
            code = compile(full, '<vba>', 'exec')
            marshal.dump(code, open(cpath, 'wb'))
        exec(code, ns)

    def call(self, name, *args):
        return self.ns[name](*args)
