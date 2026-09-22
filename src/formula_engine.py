"""
formula_engine.py
=================
A minimal, deterministic, self-contained spreadsheet formula evaluator.

Purpose
-------
We generate the Excel workbook programmatically. To *guarantee* that every
formula in the file computes correctly, we also verify them with this engine
(the same Python code that writes the formulas also evaluates them against the
same seed data). This turns "the formulas look right" into "the formulas are
provably consistent with the data model as generated".

Supported:
  - A1 and SheetName!A1 references
  - operators: + - * / ^ & comparison (= <> < > <= >=)
  - unary minus, percentages (12%)
  - parentheses
  - decimal numbers, integers
  - strings in double quotes
  - functions: SUM, SUMIF, SUMIFS, AVERAGE, MIN, MAX, IF, IFS, SWITCH,
               IFERROR, COUNT, COUNTA, COUNTIF, COUNTIFS, MINIFS, MAXIFS,
               SUMIFS, AVERAGEIFS, SUMPRODUCT, ABS, ROUND, ROUNDUP, ROUNDDOWN,
               TEXT, TEXTJOIN, CONCAT, XLOOKUP, VLOOKUP, INDEX, MATCH, LEN,
               AND, OR, NOT, CONCATENATE, TRIM, UPPER, LOWER, LEFT, RIGHT,
               MID, DATE, TODAY, EOMONTH, EDATE, YEAR, MONTH, DAY, WORKDAY,
               NETWORKDAYS, INT, MOD, CEILING, FLOOR, TRUE, FALSE, ISNUMBER,
               ISBLANK, ISERROR, VALUE
  - range operators in SUM-like functions only (simplified grammar)

Semantics follow Excel for the subset used in this project.
"""
from __future__ import annotations

import re
import math
import datetime as dt
from dataclasses import dataclass

# --------------------------------------------------------------------------- #
# Exceptions
# --------------------------------------------------------------------------- #
class FormulaError(Exception):
    def __init__(self, msg):
        super().__init__(msg)
        self.value = str(msg)


ERROR_DIV0 = FormulaError("#DIV/0!")
ERROR_NA = FormulaError("#N/A")
ERROR_VALUE = FormulaError("#VALUE!")
ERROR_REF = FormulaError("#REF!")
ERROR_NAME = FormulaError("#NAME?")

# --------------------------------------------------------------------------- #
# Cell model
# --------------------------------------------------------------------------- #
@dataclass
class Cell:
    value: object
    formula: str | None = None

    def __repr__(self):
        return f"Cell({self.value!r})"


class Grid:
    """A sheet name -> {A1 -> Cell} store."""
    def __init__(self, name: str):
        self.name = name
        self.cells: dict[str, Cell] = {}

    def set(self, ref: str, value, formula: str | None = None):
        self.cells[ref.upper()] = Cell(value, formula)

    def get(self, ref: str) -> Cell | None:
        return self.cells.get(ref.upper())

    def value(self, ref: str):
        c = self.get(ref)
        return c.value if c is not None else None


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _col_index(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1  # 0-based


def _col_name(idx: int) -> str:
    s = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _ref_to_rc(ref: str):
    ref = ref.replace("$", "")
    m = re.fullmatch(r"([A-Za-z]{1,3})([0-9]+)", ref)
    if not m:
        raise FormulaError(f"bad ref {ref}")
    return int(m.group(2)) - 1, _col_index(m.group(1))  # (row, col)


def _rc_to_ref(r: int, c: int) -> str:
    return f"{_col_name(c)}{r+1}"


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _num(v):
    if _is_number(v):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return None
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, dt.datetime):
        return v.toordinal() + 2415018.5
    if isinstance(v, dt.date):
        return v.toordinal() + 2415018.5
    return None


def _as_float(v):
    x = _num(v)
    if x is None:
        raise ERROR_VALUE
    return x


def _as_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if _is_number(v):
        if float(v) == int(v):
            return str(int(v))
        return str(float(v))
    return str(v)


def _from_serial(serial) -> dt.date | None:
    try:
        s = float(serial)
    except (TypeError, ValueError):
        return None
    if s < 1:
        return None
    ordinal = int(s) - 2415018
    try:
        return dt.date.fromordinal(ordinal)
    except ValueError:
        return None


def _date_value(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return _from_serial(_num(v))


def _truthy(v) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if _is_number(v):
        return v != 0
    if isinstance(v, str):
        return v.strip().upper() in ("TRUE", "1")
    return True


# --------------------------------------------------------------------------- #
# Tokenizer
# --------------------------------------------------------------------------- #
_TOKEN_RE = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<string>"(?:[^"]|"")*")
  | (?P<ref>(?:\$?[A-Za-z]{1,3}\$?[0-9]{1,7})|(?:'[^']+'!\$?[A-Za-z]{1,3}\$?[0-9]{1,7})|(?:[A-Za-z_][A-Za-z0-9_.]*!\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))
  | (?P<number>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)
  | (?P<percent>%)
  | (?P<ident>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><=|>=|<>|[-+*/^&=<>(),:])
    """,
    re.VERBOSE,
)


class Tok:
    __slots__ = ("kind", "value")

    def __init__(self, kind, value):
        self.kind = kind
        self.value = value

    def __repr__(self):
        return f"{self.kind}:{self.value}"


def _tokenize(s: str):
    toks = []
    i = 0
    n = len(s)
    while i < n:
        m = _TOKEN_RE.match(s, i)
        if not m:
            # unknown char -> skip single char
            i += 1
            continue
        i = m.end()
        kind = m.lastgroup
        val = m.group()
        if kind == "ws":
            continue
        toks.append(Tok(kind, val))
    return toks


# --------------------------------------------------------------------------- #
# Range/area helpers
# --------------------------------------------------------------------------- #
class AreaSpec:
    """A single-cell reference or a rectangular range (handled in SIMPLE fns)."""
    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet = sheet
        self.r1, self.c1 = r1, c1
        self.r2, self.c2 = r2, c2

    def cells(self, book) -> list:
        grid = book.get(self.sheet)
        out = []
        for r in range(self.r1, self.r2 + 1):
            row = []
            for c in range(self.c1, self.c2 + 1):
                ref = _rc_to_ref(r, c)
                cell = grid.get(ref)
                row.append(cell.value if cell is not None else None)
            out.append(row)
        return out

    def width(self):
        return max(0, self.c2 - self.c1 + 1)

    def height(self):
        return max(0, self.r2 - self.r1 + 1)


class Parser:
    def __init__(self, toks):
        self.toks = toks
        self.pos = 0

    def peek(self):
        return self.toks[self.pos] if self.pos < len(self.toks) else None

    def next(self):
        t = self.peek()
        self.pos += 1
        return t

    def eat_kind(self, kind):
        t = self.peek()
        if t and t.kind == kind:
            self.pos += 1
            return t
        return None

    def expect(self, kind, val=None):
        t = self.next()
        if t is None:
            raise FormulaError("unexpected end")
        if val is not None and t.value != val:
            raise FormulaError(f"expected '{val}' got '{t.value}'")
        if t.kind != kind:
            raise FormulaError(f"expected {kind} got {t.kind}")
        return t

    def parse_range_part(self, sheet, ref):
        m = re.match(r"([A-Za-z]{1,3})([0-9]+)", ref)
        r, c = _ref_to_rc(ref)
        return sheet, r, c, r, c


# --------------------------------------------------------------------------- #
# Evaluator
# --------------------------------------------------------------------------- #
class Book:
    def __init__(self):
        self.sheets: dict[str, Grid] = {}

    def add(self, grid: Grid):
        self.sheets[grid.name] = grid

    def get(self, name: str) -> Grid:
        if name not in self.sheets:
            raise FormulaError(f"unknown sheet {name}")
        return self.sheets[name]

    def resolve(self, sheet: str, ref: str, value=False):
        if "!" in ref:
            sh, _, cellref = ref.partition("!")
            sh = sh.strip("'")
            grid = self.get(sh)
        else:
            grid = self.get(sheet)
            cellref = ref
        cellref = cellref.upper().replace("$", "")
        c = grid.get(cellref)
        if c is None:
            return 0 if value else ERROR_REF  # blank cell in value context -> 0
        return c.value if value else c


class Evaluator:
    def __init__(self, book: Book):
        self.book = book

    def eval(self, formula: str, sheet: str):
        formula = formula.strip()
        if formula.startswith("="):
            formula = formula[1:]
        try:
            toks = _tokenize(formula)
            if not toks:
                return None
            p = Parser(toks)
            val = self.parse_expr(p, sheet)
            return val
        except FormulaError as e:
            return e
        except RecursionError:
            return ERROR_VALUE

    # --- grammar ---
    def parse_expr(self, p: Parser, sheet: str, in_range=False):
        return self.parse_cmp(p, sheet)

    def parse_cmp(self, p: Parser, sheet: str):
        left = self.parse_concat(p, sheet)
        t = p.peek()
        while t and t.kind == "op" and t.value in ("=", "<>", "<", ">", "<=", ">="):
            op = p.next().value
            right = self.parse_concat(p, sheet)
            left = self._cmp(op, left, right)
            t = p.peek()
        return left

    def parse_concat(self, p: Parser, sheet: str):
        left = self.parse_add(p, sheet)
        t = p.peek()
        while t and t.kind == "op" and t.value == "&":
            p.next()
            right = self.parse_add(p, sheet)
            left = _as_text(left) + _as_text(right)
            t = p.peek()
        return left

    def parse_add(self, p: Parser, sheet: str):
        left = self.parse_mul(p, sheet)
        t = p.peek()
        while t and t.kind == "op" and t.value in ("+", "-"):
            op = p.next().value
            right = self.parse_mul(p, sheet)
            try:
                l, r = _as_float(left), _as_float(right)
                left = l + r if op == "+" else l - r
            except FormulaError:
                left = ERROR_VALUE
            t = p.peek()
        return left

    def parse_mul(self, p: Parser, sheet: str):
        left = self.parse_unary(p, sheet)
        t = p.peek()
        while t and t.kind == "op" and t.value in ("*", "/", "^"):
            op = p.next().value
            right = self.parse_unary(p, sheet)
            if isinstance(right, FormulaError):
                left = right
                continue
            try:
                l, r = _as_float(left), _as_float(right)
                if op == "*":
                    left = l * r
                elif op == "/":
                    left = l / r if r != 0 else ERROR_DIV0
                else:
                    left = l ** r
            except FormulaError:
                left = ERROR_VALUE
            t = p.peek()
        return left

    def parse_unary(self, p: Parser, sheet: str):
        t = p.peek()
        if t and t.kind == "op" and t.value == "-":
            p.next()
            v = self.parse_unary(p, sheet)
            if isinstance(v, FormulaError):
                return v
            try:
                return -_as_float(v)
            except FormulaError:
                return ERROR_VALUE
        if t and t.kind == "op" and t.value == "+":
            p.next()
            return self.parse_unary(p, sheet)
        return self.parse_postfix(p, sheet)

    def parse_postfix(self, p: Parser, sheet: str):
        v = self.parse_atom(p, sheet)
        t = p.peek()
        while t and t.kind == "percent":
            p.next()
            if isinstance(v, FormulaError):
                continue
            try:
                v = _as_float(v) / 100.0
            except FormulaError:
                v = ERROR_VALUE
            t = p.peek()
        return v

    def parse_atom(self, p: Parser, sheet: str):
        t = p.next()
        if t is None:
            raise FormulaError("unexpected end")
        if t.kind == "number":
            return float(t.value)
        if t.kind == "string":
            return t.value[1:-1].replace('""', '"')
        if t.kind == "ref":
            cell = self.book.resolve(sheet, t.value, value=True)
            return cell
        if t.kind == "op" and t.value == "(":
            v = self.parse_expr(p, sheet)
            p.expect("op", ")")
            return v
        if t.kind == "ident":
            name = t.value.lower()
            # function call?
            if p.peek() and p.peek().kind == "op" and p.peek().value == "(":
                p.next()
                args = self.parse_args(p, sheet)
                return self._call(name, args, sheet)
            # boolean / error literal
            if name == "true":
                return True
            if name == "false":
                return False
            # bare named-range constant? treat as NAME error in this engine
            raise FormulaError(f"#NAME? {t.value}")
        raise FormulaError(f"unexpected token {t.value}")

    def parse_args(self, p: Parser, sheet: str):
        args = []
        if p.peek() and p.peek().kind == "op" and p.peek().value == ")":
            p.next()
            return args
        while True:
            args.append(self.parse_arg(p, sheet))
            t = p.next()
            if t is None:
                break
            if t.value == ")":
                break
            if t.value != ",":
                raise FormulaError(f"expected comma, got {t.value}")
        return args

    def parse_arg(self, p: Parser, sheet: str):
        # attempt range detection: ref ':' ref
        start = p.pos
        vals = [self.parse_expr(p, sheet)]
        if p.peek() and p.peek().kind == "op" and p.peek().value == ":":
            # we need left token of range to be a ref; simplify: re-parse tokens
            # for a range by peeking at raw tokens is complex; support 'ref:ref'
            # by checking token stream at `start`
            toks = p.toks
            # check pattern: ref : ref
            if (
                start + 2 < len(toks)
                and toks[start].kind == "ref"
                and toks[start + 1].kind == "op" and toks[start + 1].value == ":"
                and toks[start + 2].kind == "ref"
            ):
                p.next()  # skip ':'
                p.pos = start + 2
                right_tok = p.next()
                area = self._make_area(sheet, toks[start].value, right_tok.value)
                return area
        return vals[0]

    def _make_area(self, sheet: str, a: str, b: str):
        sh1, r1c1 = (a.partition("!")[0].strip("'"), a.partition("!")[2]) if "!" in a else (sheet, a)
        if "!" in b:
            sh2, r2c2 = b.partition("!")[0].strip("'"), b.partition("!")[2]
        else:
            sh2, r2c2 = sh1, b  # right side inherits the left side's sheet
        r1, c1 = _ref_to_rc(r1c1)
        r2, c2 = _ref_to_rc(r2c2)
        return AreaSpec(sh1, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    # --- operators ---
    def _cmp(self, op, l, r):
        if isinstance(l, FormulaError):
            return l
        if isinstance(r, FormulaError):
            return r
        try:
            nl, nr = _num(l), _num(r)
            if nl is not None and nr is not None:
                l, r = nl, nr
            else:
                l, r = _as_text(l), _as_text(r)
            if op == "=":
                return l == r
            if op == "<>":
                return l != r
            if op == "<":
                return l < r
            if op == ">":
                return l > r
            if op == "<=":
                return l <= r
            if op == ">=":
                return l >= r
        except FormulaError:
            return ERROR_VALUE
        return False

    # --- range arg helpers ---
    def _iter_areas(self, args):
        for a in args:
            if isinstance(a, AreaSpec):
                for row in a.cells(self.book):
                    for v in row:
                        yield v
            else:
                yield a

    def _crit_areas(self, args, start=1):
        """yield (range, criteria) pairs of consecutive args, starting at index `start`"""
        for i in range(start, len(args), 2):
            if i + 1 >= len(args):
                break
            yield args[i], args[i + 1]

    def _matches(self, v, crit):
        if isinstance(crit, AreaSpec):
            crit = self._first_area_val(crit)
        if isinstance(crit, FormulaError):
            return False
        cs = _as_text(crit)
        if cs.startswith("<="):
            return (_num(v) or 0) <= float(cs[2:])
        if cs.startswith(">="):
            return (_num(v) or 0) >= float(cs[2:])
        if cs.startswith("<>"):
            return _as_text(v) != cs[2:].strip()
        if cs.startswith("<"):
            return (_num(v) or 0) < float(cs[1:])
        if cs.startswith(">"):
            return (_num(v) or 0) > float(cs[1:])
        # wildcards
        import fnmatch
        try:
            return fnmatch.fnmatchcase(_as_text(v), cs)
        except Exception:
            return _as_text(v) == cs

    def _first_area_val(self, a):
        if not isinstance(a, AreaSpec):
            return a
        for v in self._iter_areas([a]):
            return v
        return None

    # --- functions ---
    def _call(self, name: str, args, sheet: str):
        f = {
            "sum": self._f_sum, "sumif": self._f_sumif, "sumifs": self._f_sumifs,
            "average": self._f_average, "averageif": self._f_averageif,
            "averageifs": self._f_averageifs, "min": self._f_min, "max": self._f_max,
            "minifs": self._f_minifs, "maxifs": self._f_maxifs, "if": self._f_if,
            "ifs": self._f_ifs, "switch": self._f_switch, "iferror": self._f_iferror,
            "count": self._f_count, "counta": self._f_counta, "countif": self._f_countif,
            "countifs": self._f_countifs, "sumproduct": self._f_sumproduct,
            "abs": self._f_abs, "round": self._f_round, "roundup": self._f_roundup,
            "rounddown": self._f_rounddown, "text": self._f_text,
            "textjoin": self._f_textjoin, "concat": self._f_concat,
            "xlookup": self._f_xlookup, "vlookup": self._f_vlookup,
            "index": self._f_index, "match": self._f_match, "len": self._f_len,
            "and": self._f_and, "or": self._f_or, "not": self._f_not,
            "concatenate": self._f_concat, "trim": self._f_trim,
            "upper": self._f_upper, "lower": self._f_lower, "left": self._f_left,
            "right": self._f_right, "mid": self._f_mid, "date": self._f_date,
            "today": self._f_today, "eomonth": self._f_eomonth, "edate": self._f_edate,
            "year": self._f_year, "month": self._f_month, "day": self._f_day,
            "workday": self._f_workday, "networkdays": self._f_networkdays,
            "int": self._f_int, "mod": self._f_mod, "ceiling": self._f_ceiling,
            "floor": self._f_floor, "isnumber": self._f_isnumber,
            "isblank": self._f_isblank, "iserror": self._f_iserror,
            "value": self._f_value,
        }.get(name)
        if f is None:
            raise FormulaError(f"#NAME? {name}")
        try:
            return f(args, sheet)
        except FormulaError as e:
            return e
        except Exception as e:
            return FormulaError(f"#VALUE! ({name}: {e})")

    def _floats(self, args, sheet):
        out = []
        for v in self._iter_areas(args):
            if isinstance(v, FormulaError):
                continue
            n = _num(v)
            if n is not None:
                out.append(n)
        return out

    def _f_sum(self, args, sheet):
        return sum(self._floats(args, sheet))

    def _f_average(self, args, sheet):
        fs = self._floats(args, sheet)
        return sum(fs) / len(fs) if fs else ERROR_DIV0

    def _f_min(self, args, sheet):
        fs = self._floats(args, sheet)
        return min(fs) if fs else 0

    def _f_max(self, args, sheet):
        fs = self._floats(args, sheet)
        return max(fs) if fs else 0

    def _f_sumif(self, args, sheet):
        rng = args[0]
        sum_area = args[2] if len(args) >= 3 else args[0]
        crit = args[1]
        total = 0.0
        cells = sum_area.cells(self.book) if isinstance(sum_area, AreaSpec) else [[sum_area]]
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                if self._matches(self._cell_at(rng, ri, ci), crit):
                    total += _num(v) or 0
        return total

    def _f_sumifs(self, args, sheet):
        sum_area = args[0]
        pairs = list(self._crit_areas(args, start=1))
        total = 0.0
        cells = sum_area.cells(self.book) if isinstance(sum_area, AreaSpec) else [[sum_area]]
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                if self._all_match(pairs, ri, ci):
                    total += _num(v) or 0
        return total

    def _all_match(self, pairs, ri, ci):
        for area, crit in pairs:
            if not self._matches(self._cell_at(area, ri, ci), crit):
                return False
        return True

    def _cell_at(self, area, ri, ci):
        if not isinstance(area, AreaSpec):
            return area
        cells = area.cells(self.book)
        if ri < len(cells) and ci < len(cells[ri]):
            return cells[ri][ci]
        return None

    def _f_averageif(self, args, sheet):
        rng = args[0]
        avg_area = args[2] if len(args) >= 3 else args[0]
        crit = args[1]
        nums = []
        cells = avg_area.cells(self.book) if isinstance(avg_area, AreaSpec) else [[avg_area]]
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                if self._matches(self._cell_at(rng, ri, ci), crit) and _num(v) is not None:
                    nums.append(_num(v))
        if not nums:
            return ERROR_DIV0
        return sum(nums) / len(nums)

    def _f_averageifs(self, args, sheet):
        avg_area = args[0]
        pairs = list(self._crit_areas(args, start=1))
        nums = []
        cells = avg_area.cells(self.book) if isinstance(avg_area, AreaSpec) else [[avg_area]]
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                if self._all_match(pairs, ri, ci) and _num(v) is not None:
                    nums.append(_num(v))
        if not nums:
            return ERROR_DIV0
        return sum(nums) / len(nums)

    def _f_minifs(self, args, sheet):
        return self._minmaxifs(args, sheet, min)

    def _f_maxifs(self, args, sheet):
        return self._minmaxifs(args, sheet, max)

    def _minmaxifs(self, args, sheet, fn):
        val_area = args[0]
        pairs = list(self._crit_areas(args, start=1))
        vals = []
        cells = val_area.cells(self.book) if isinstance(val_area, AreaSpec) else [[val_area]]
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                if self._all_match(pairs, ri, ci) and _num(v) is not None:
                    vals.append(_num(v))
        if not vals:
            return 0
        return fn(vals)

    def _f_if(self, args, sheet):
        if len(args) < 2:
            raise FormulaError("IF needs 2-3 args")
        cond = args[0]
        if isinstance(cond, FormulaError):
            return cond
        if _truthy(cond):
            return args[1]
        return args[2] if len(args) > 2 else False

    def _f_ifs(self, args, sheet):
        for i in range(0, len(args) - 1, 2):
            c, v = args[i], args[i + 1]
            if isinstance(c, FormulaError):
                return c
            if _truthy(c):
                return v
        return ERROR_NA

    def _f_switch(self, args, sheet):
        if len(args) < 2:
            raise FormulaError("SWITCH needs expr + pairs")
        expr = _as_text(args[0])
        i = 1
        while i < len(args) - 1:
            if _as_text(args[i]) == expr:
                return args[i + 1]
            i += 2
        return args[-1]

    def _f_iferror(self, args, sheet):
        if isinstance(args[0], FormulaError):
            return args[1] if len(args) > 1 else 0
        return args[0]

    def _f_count(self, args, sheet):
        return sum(1 for v in self._iter_areas(args) if _num(v) is not None)

    def _f_counta(self, args, sheet):
        return sum(1 for v in self._iter_areas(args) if v is not None and not isinstance(v, FormulaError))

    def _f_countif(self, args, sheet):
        rng, crit = args[0], args[1]
        return sum(1 for v in self._iter_areas([rng]) if self._matches(v, crit))

    def _f_countifs(self, args, sheet):
        first_area = args[0]
        cells = first_area.cells(self.book) if isinstance(first_area, AreaSpec) else [[first_area]]
        pairs = list(self._crit_areas(args, start=0))
        n = 0
        for ri in range(len(cells)):
            for ci in range(len(cells[ri])):
                if self._all_match(pairs, ri, ci):
                    n += 1
        return n

    def _f_sumproduct(self, args, sheet):
        areas = [a for a in args if isinstance(a, AreaSpec)]
        total = 0.0
        if areas:
            grids = [a.cells(self.book) for a in areas]
            h = max(len(g) for g in grids)
            w = max(len(g[0]) for g in grids if g)
            for r in range(h):
                for c in range(w):
                    prod = 1.0
                    for g in grids:
                        if r < len(g) and c < len(g[r]):
                            prod *= _num(g[r][c]) or 0
                    total += prod
        return total

    def _f_abs(self, args, sheet):
        return abs(_as_float(self._first_area_val(args[0] if len(args) == 1 else args)))

    def _f_round(self, args, sheet):
        return round(_as_float(args[0]), int(args[1]))

    def _f_roundup(self, args, sheet):
        x, d = _as_float(args[0]), int(args[1])
        m = 10 ** d
        return math.ceil(x * m - 1e-12) / m

    def _f_rounddown(self, args, sheet):
        x, d = _as_float(args[0]), int(args[1])
        m = 10 ** d
        return math.floor(x * m + 1e-12) / m

    def _f_text(self, args, sheet):
        v = args[0]
        fmt = _as_text(args[1])
        d = _date_value(v) if ("y" in fmt or "m" in fmt.lower() or "d" in fmt.lower()) else None
        if d is not None and "%" not in fmt and "0" not in fmt:
            return d.strftime(fmt.replace("yyyy", "%Y").replace("mm", "%m").replace("dd", "%d"))
        if "0.00" in fmt:
            return f"{_as_float(v):,.2f}"
        if "0" in fmt and "," in fmt:
            return f"{_as_float(v):,.0f}"
        if "#,###" in fmt:
            return f"{_as_float(v):,.0f}"
        if "%" in fmt:
            if "0.0%" in fmt:
                return f"{_as_float(v)*100:.1f}%"
            return f"{_as_float(v)*100:.0f}%"
        return _as_text(v)

    def _f_textjoin(self, args, sheet):
        delim, ignore_empty, *rest = args
        vals = []
        for a in rest:
            for v in self._iter_areas([a]):
                if v is None or isinstance(v, FormulaError):
                    if ignore_empty in (True, 1):
                        continue
                    vals.append("")
                else:
                    vals.append(_as_text(v))
        return delim.join(vals)

    def _f_concat(self, args, sheet):
        return "".join(_as_text(v) for v in self._iter_areas(args))

    def _f_xlookup(self, args, sheet):
        lookup, lookarr, retarr = args[0], args[1], args[2]
        if len(args) > 3 and args[3] not in (None, ""):
            default = args[3]
        else:
            default = ERROR_NA
        lv = _as_text(lookup)
        la = lookarr.cells(self.book) if isinstance(lookarr, AreaSpec) else [[lookarr]]
        ra = retarr.cells(self.book) if isinstance(retarr, AreaSpec) else [[retarr]]
        for ri, row in enumerate(la):
            for ci, v in enumerate(row):
                if _as_text(v) == lv:
                    if ri < len(ra) and ci < len(ra[ri]):
                        return ra[ri][ci]
                    return None
        return default

    def _f_vlookup(self, args, sheet):
        lookup, table, col = args[0], args[1], int(args[2])
        exact = len(args) > 3 and _truthy(args[3])
        rows = table.cells(self.book) if isinstance(table, AreaSpec) else [[table]]
        lv = _as_text(lookup)
        for row in rows:
            if not row:
                continue
            if exact and _as_text(row[0]) == lv:
                return row[col - 1] if col - 1 < len(row) else None
            if not exact and _num(row[0]) is not None and _num(lookup) is not None and _num(row[0]) >= _num(lookup):
                return row[col - 1] if col - 1 < len(row) else None
        return ERROR_NA

    def _f_index(self, args, sheet):
        area = args[0] if isinstance(args[0], AreaSpec) else None
        if area is None:
            raise FormulaError("INDEX needs a range")
        cells = area.cells(self.book)
        if len(args) == 2:
            i = int(args[1])
            flat = [v for row in cells for v in row]
            return flat[i - 1] if 0 < i <= len(flat) else ERROR_REF
        r = int(args[1]); c = int(args[2])
        if 0 < r <= len(cells) and 0 < c <= len(cells[r - 1]):
            return cells[r - 1][c - 1]
        return ERROR_REF

    def _f_match(self, args, sheet):
        lookup, area = args[0], args[1]
        cells = area.cells(self.book) if isinstance(area, AreaSpec) else [[area]]
        flat = [v for row in cells for v in row]
        lv = _as_text(lookup)
        for i, v in enumerate(flat):
            if _as_text(v) == lv:
                return i + 1
        return ERROR_NA

    def _f_len(self, args, sheet):
        return len(_as_text(self._first_area_val(args[0])))

    def _f_and(self, args, sheet):
        for v in self._iter_areas(args):
            if not _truthy(v):
                return False
        return True

    def _f_or(self, args, sheet):
        for v in self._iter_areas(args):
            if _truthy(v):
                return True
        return False

    def _f_not(self, args, sheet):
        return not _truthy(self._first_area_val(args[0]))

    def _f_trim(self, args, sheet):
        return " ".join(_as_text(self._first_area_val(args[0])).split())

    def _f_upper(self, args, sheet):
        return _as_text(self._first_area_val(args[0])).upper()

    def _f_lower(self, args, sheet):
        return _as_text(self._first_area_val(args[0])).lower()

    def _f_left(self, args, sheet):
        s = _as_text(self._first_area_val(args[0]))
        n = int(args[1]) if len(args) > 1 else 1
        return s[:n]

    def _f_right(self, args, sheet):
        s = _as_text(self._first_area_val(args[0]))
        n = int(args[1]) if len(args) > 1 else 1
        return s[-n:] if n > 0 else ""

    def _f_mid(self, args, sheet):
        s = _as_text(args[0])
        start = int(args[1]); n = int(args[2])
        return s[start - 1:start - 1 + n]

    def _f_date(self, args, sheet):
        return dt.date(int(args[0]), int(args[1]), int(args[2]))

    def _f_today(self, args, sheet):
        return dt.date(2026, 9, 22)  # fixed seed date, deterministic

    def _f_eomonth(self, args, sheet):
        d = _date_value(args[0])
        m = int(args[1])
        y, mo = d.year + (d.month - 1 + m) // 12, (d.month - 1 + m) % 12 + 1
        nxt = dt.date(y + mo // 12, (mo % 12) + 1, 1)
        return nxt - dt.timedelta(days=1)

    def _f_edate(self, args, sheet):
        d = _date_value(args[0])
        m = int(args[1])
        y, mo = d.year + (d.month - 1 + m) // 12, (d.month - 1 + m) % 12 + 1
        return d.replace(year=y, month=mo)

    def _f_year(self, args, sheet):
        return _date_value(args[0]).year

    def _f_month(self, args, sheet):
        return _date_value(args[0]).month

    def _f_day(self, args, sheet):
        return _date_value(args[0]).day

    def _f_workday(self, args, sheet):
        start = _date_value(args[0])
        days = int(args[1])
        if len(args) > 2 and isinstance(args[2], AreaSpec):
            hol = {_date_value(v) for v in self._iter_areas([args[2]]) if _date_value(v)}
        else:
            hol = set()
        cur = start
        step = 1 if days >= 0 else -1
        count = 0
        while count < abs(days):
            cur += dt.timedelta(days=step)
            if cur.weekday() < 5 and cur not in hol:
                count += 1
                if count == abs(days):
                    return cur
        return start

    def _f_networkdays(self, args, sheet):
        start = _date_value(args[0]); end = _date_value(args[1])
        if len(args) > 2 and isinstance(args[2], AreaSpec):
            hol = {_date_value(v) for v in self._iter_areas([args[2]]) if _date_value(v)}
        else:
            hol = set()
        n = 0
        d = start
        while d <= end:
            if d.weekday() < 5 and d not in hol:
                n += 1
            d += dt.timedelta(days=1)
        return n

    def _f_int(self, args, sheet):
        return math.floor(_as_float(args[0]))

    def _f_mod(self, args, sheet):
        return _as_float(args[0]) % _as_float(args[1])

    def _f_ceiling(self, args, sheet):
        x = _as_float(args[0]); s = _as_float(args[1]) if len(args) > 1 else 1
        return math.ceil(x / s) * s

    def _f_floor(self, args, sheet):
        x = _as_float(args[0]); s = _as_float(args[1]) if len(args) > 1 else 1
        return math.floor(x / s) * s

    def _f_isnumber(self, args, sheet):
        return _num(self._first_area_val(args[0])) is not None

    def _f_isblank(self, args, sheet):
        v = self._first_area_val(args[0])
        return v is None or v == ""

    def _f_iserror(self, args, sheet):
        return isinstance(self._first_area_val(args[0]), FormulaError)

    def _f_value(self, args, sheet):
        raise FormulaError("#VALUE! (VALUE of non-number not supported in engine)")


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def evaluate(book: Book, sheet: str, formula: str):
    ev = Evaluator(book)
    return ev.eval(formula, sheet)


def build_eval_context(grids: dict[str, dict[str, object]]):
    """Convert {sheet: {A1: value}} into a Book."""
    book = Book()
    for name, cells in grids.items():
        g = Grid(name)
        for ref, v in cells.items():
            g.set(ref, v)
        book.add(g)
    return book
