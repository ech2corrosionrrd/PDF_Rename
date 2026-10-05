"""Модель аркуша «база» книги «Лічильники по підстанціям» (без Tk і без Excel).

Повторює логіку VBA-форми `frmBase`/`modBase`: поля будуються за заголовками
(рядки 1–3), групуються за об'єднаним заголовком 1-го рівня, сусідні стовпці
«Квартал / Рік», «Число / Місяць / Рік» та «№ / А, В, С» виводяться одним блоком.
Читання з файлу — `meters_excel.py`, вікно — `meters_form.py`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple, Union

SH_BASE = "база"
SH_REF = "Довідники"
SH_HISTORY = "Історія_Змін"

HDR_TOP = 1
HDR_BOTTOM = 3
FIRST_DATA_ROW = 4
LAST_COL = 60  # стовпці A..BH

COL_NUM = 1
COL_EIC = 2
COL_SUBST = 3
COL_CONN = 5
COL_METER = 7
ID_MIN = 5
MAX_LIST = 30
ALL_SUBSTATIONS = "(усі підстанції)"

REF_KEY_ROW = 4
REF_FIRST_ROW = 6

# Підписи, що перекривають заголовок аркуша (як у VBA-формі)
CAPTION_OVERRIDES = {56: "Код виконавця", 57: "Номер акту"}
_GENERAL_GROUP = "Загальні"
_COL_MAJOR_GROUPS = ("Лічильник",)
_COL_MAJOR_PREFIXES = ("Основні дані трансформаторів",)

Value = Union[None, str, int, float, date, datetime]


def col_letter(c: int) -> str:
    s = ""
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


# ---------------------------------------------------------------- знімок аркуша
@dataclass
class Snapshot:
    """Вміст аркуша на момент читання. Індекси рядків — як на аркуші (з 1)."""

    # hdr[i][c-1]: текст шапки (i=0..2), для об'єднаних комірок — значення якоря
    hdr: List[List[str]]
    # row1_wide[c-1]: стовпець c входить в об'єднання 1-го рядка шириною > 1
    row1_wide: List[bool]
    # values[r - FIRST_DATA_ROW][c-1]: текст комірки (для об'єднаних — якоря)
    values: List[List[str]]
    # merge_anchor[(r, c)] = рядок-якір, якщо комірка входить в об'єднання по вертикалі
    merge_anchor: Dict[Tuple[int, int], int] = field(default_factory=dict)
    # довідники: ключ -> значення
    ref_lists: Dict[str, List[str]] = field(default_factory=dict)

    @property
    def last_row(self) -> int:
        return FIRST_DATA_ROW + len(self.values) - 1

    def rows(self) -> range:
        return range(FIRST_DATA_ROW, self.last_row + 1)

    def text(self, r: int, c: int) -> str:
        i = r - FIRST_DATA_ROW
        if i < 0 or i >= len(self.values):
            return ""
        return self.values[i][c - 1]

    def set_text(self, r: int, c: int, s: str) -> None:
        """Оновлює кеш після запису (для об'єднаних — усі комірки блоку)."""
        anchor = self.merge_anchor.get((r, c), r)
        for rr in self.rows():
            if self.merge_anchor.get((rr, c), rr) == anchor:
                self.values[rr - FIRST_DATA_ROW][c - 1] = s

    def subst_of(self, r: int) -> str:
        return self.text(r, COL_SUBST).strip()

    def is_locked(self, r: int, c: int) -> bool:
        """Колонка підстанції або не-перший рядок об'єднаного блоку — лише читання."""
        if c == COL_SUBST:
            return True
        return self.merge_anchor.get((r, c), r) != r

    def substations(self) -> List[str]:
        seen: Dict[str, None] = {}
        for r in self.rows():
            s = self.subst_of(r)
            if s:
                seen.setdefault(s, None)
        return list(seen)

    def last_row_of_subst(self, subst: str) -> int:
        last = 0
        for r in self.rows():
            if self.subst_of(r) == subst:
                last = r
        return last

    def connections(self, subst: str) -> List[str]:
        seen: Dict[str, None] = {}
        for r in self.rows():
            if self.subst_of(r) == subst:
                s = self.text(r, COL_CONN).strip()
                if s:
                    seen.setdefault(s, None)
        return list(seen)


# ---------------------------------------------------------------- заголовки
def group_of(snap: Snapshot, c: int) -> str:
    if snap.row1_wide[c - 1]:
        s = snap.hdr[0][c - 1].strip()
        if s:
            return s
    return _GENERAL_GROUP


def caption_of(snap: Snapshot, c: int) -> str:
    if c in CAPTION_OVERRIDES:
        return CAPTION_OVERRIDES[c]
    grp = group_of(snap, c)
    parts: List[str] = []
    prev = ""
    for i in range(3):
        part = snap.hdr[i][c - 1].strip()
        if part and part != grp and part != prev:
            parts.append(part)
            prev = part
    return " / ".join(parts) if parts else "Стовпець " + col_letter(c)


# ---------------------------------------------------------------- розкладка форми
@dataclass
class Block:
    """Одне поле форми з 1–3 стовпців (дата, квартал/рік, фази)."""

    cols: Tuple[int, ...]
    label: str
    sep: str = ""  # роздільник частин при друці: " / ", ".", "#" (позиційний)
    roles: Tuple[str, ...] = ()  # "Q", "Y", "D", "M" — для списків значень
    captions: Tuple[str, ...] = ()  # підказки по кожному стовпцю


@dataclass
class Group:
    name: str
    blocks: List[Block] = field(default_factory=list)
    column_major: bool = False


def _field_rank(c: int) -> int:
    ranks = {1: 1, COL_CONN: 2, 2: 3, 55: 4, 57: 1, 58: 2, 59: 3, 60: 4, 56: 5}
    return ranks.get(c, 100 + c)


def build_layout(snap: Snapshot) -> List[Group]:
    caps = {c: caption_of(snap, c) for c in range(1, LAST_COL + 1)}
    grps = {c: group_of(snap, c) for c in range(1, LAST_COL + 1)}
    order: List[str] = []
    for c in range(1, LAST_COL + 1):
        if grps[c] not in order:
            order.append(grps[c])
    # стабільне сортування: група, потім ранг поля
    cols = sorted(
        range(1, LAST_COL + 1),
        key=lambda c: (order.index(grps[c]) * 1000 + _field_rank(c), c),
    )

    nxt: Dict[int, int] = {}
    second: Set[int] = set()
    info: Dict[int, Block] = {}

    def link(chain: Sequence[int], label: str, sep: str, roles: Tuple[str, ...]) -> None:
        for a, b in zip(chain, chain[1:]):
            nxt[a] = b
        second.update(chain[1:])
        info[chain[0]] = Block(
            cols=tuple(chain),
            label=label,
            sep=sep,
            roles=roles,
            captions=tuple(caps[x] for x in chain),
        )

    n = len(cols)
    for i in range(n - 1):  # Квартал + Рік
        c, d = cols[i], cols[i + 1]
        if grps[c] == grps[d] and caps[c].endswith(" / Квартал"):
            pre = caps[c][: -len(" / Квартал")]
            if caps[d] == pre + " / Рік":
                link((c, d), pre + " (квартал / рік)", " / ", ("Q", "Y"))
    for i in range(n - 2):  # Число + Місяць + Рік
        c, d, e = cols[i], cols[i + 1], cols[i + 2]
        if grps[c] == grps[d] == grps[e] and caps[c].endswith(" / Число"):
            pre = caps[c][: -len(" / Число")]
            if caps[d] == pre + " / Місяць" and caps[e] == pre + " / Рік":
                link((c, d, e), pre + " (число / місяць / рік)", ".", ("D", "M", "Y"))
    for i in range(n - 2):  # № за фазами
        c, d, e = cols[i], cols[i + 1], cols[i + 2]
        if grps[c] == grps[d] == grps[e] and all(
            len(caps[x]) == 5 and caps[x].startswith("№ / ") for x in (c, d, e)
        ):
            if c not in second and c not in nxt:
                label = "№ ({} / {} / {})".format(caps[c][4:], caps[d][4:], caps[e][4:])
                link((c, d, e), label, "#", ())
    for i in range(n - 2):  # Дата акта
        c, d, e = cols[i], cols[i + 1], cols[i + 2]
        if grps[c] == grps[d] == grps[e] and caps[c] == "Число":
            if caps[d] == "Місяць" and caps[e] == "Рік":
                link((c, d, e), "Дата акта", ".", ("D", "M", "Y"))

    layout: Dict[str, Group] = {
        g: Group(
            g,
            column_major=g in _COL_MAJOR_GROUPS
            or any(g.startswith(p) for p in _COL_MAJOR_PREFIXES),
        )
        for g in order
    }
    seen: Set[str] = set()
    for c in cols:
        if c in second:
            continue
        blk = info.get(c)
        if blk is None:
            cap = caps[c]
            if cap in seen:
                cap = "{} ({})".format(cap, col_letter(c))
            else:
                seen.add(cap)
            blk = Block(cols=(c,), label=cap, captions=(caps[c],))
        layout[grps[c]].blocks.append(blk)
    return [layout[g] for g in order]


# ---------------------------------------------------------------- списки значень
def date_list(role: str) -> List[str]:
    if role == "Q":
        return ["I", "II", "III", "IV"]
    if role == "D":
        return [str(i) for i in range(1, 32)]
    if role == "M":
        return [str(i) for i in range(1, 13)]
    return [str(i) for i in range(2010, date.today().year + 11)]


_ROLE_KEYS = {"Q": "Квартал", "Y": "Рік", "D": "Число", "M": "Місяць"}


def build_value_lists(snap: Snapshot) -> Dict[int, List[str]]:
    """Значення, що вже є у стовпці (не більше MAX_LIST різних) — для випадаючих списків."""
    out: Dict[int, List[str]] = {}
    for c in range(1, LAST_COL + 1):
        seen: Dict[str, None] = {}
        for r in snap.rows():
            s = snap.text(r, c).strip()
            if s:
                seen.setdefault(s, None)
                if len(seen) > MAX_LIST:
                    break
        if 0 < len(seen) <= MAX_LIST:
            out[c] = sorted(seen, key=str.casefold)
    return out


def roles_by_col(layout: Iterable[Group]) -> Dict[int, str]:
    out: Dict[int, str] = {}
    for g in layout:
        for b in g.blocks:
            for c, role in zip(b.cols, b.roles):
                out[c] = role
    return out


def list_for(
    c: int,
    roles: Dict[int, str],
    ref: Dict[str, List[str]],
    auto: Dict[int, List[str]],
) -> Optional[List[str]]:
    """Список значень для поля: довідник > дата-список > значення таблиці; None — вільний ввід."""
    role = roles.get(c)
    if role:
        return list(ref.get(_ROLE_KEYS[role]) or date_list(role))
    if col_letter(c) in ref:
        return list(ref[col_letter(c)])
    if c in auto:
        return list(auto[c])
    return None


# ---------------------------------------------------------------- значення комірок
def format_value(v: object) -> str:
    """Текст для поля форми (як CStr/Format у VBA, з комою як десятковим роздільником)."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return repr(v).replace(".", ",")
    return str(v).strip() if isinstance(v, str) else str(v)


_NUM_RE = re.compile(r"-?\d+(?:[.,]\d+)?")
_DATE_RE = re.compile(r"(\d{1,2})[./](\d{1,2})[./](\d{4})")


def parse_cell_text(s: str) -> Tuple[Value, bool]:
    """Текст з форми → (значення, примусово_текст).

    None — очистити комірку. Нулі на початку («0999») та рядки з «=» зберігаються текстом,
    щоб Excel не з'їв їх і не перетворив на формулу.
    """
    s = (s or "").strip()
    if not s:
        return None, False
    if len(s) > 1 and s[0] == "0" and s.isdigit():
        return s, True
    if s[0] == "=":
        return s, True
    if _NUM_RE.fullmatch(s):
        if "," not in s and "." not in s:
            return int(s), False
        return float(s.replace(",", ".")), False
    m = _DATE_RE.fullmatch(s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))), False
        except ValueError:
            pass
    return s, False


# ---------------------------------------------------------------- пошук
def norm_id(s: str) -> str:
    s = (s or "").strip().upper()
    for ch in (" ", " ", "-", ".", "_"):
        s = s.replace(ch, "")
    return s


_SPLIT_RE = re.compile(r"[\r\n,;/\s]+")


def parse_id_query(text: str) -> Tuple[List[str], Optional[str]]:
    """Розбір рядка пошуку за EIC / № лічильника → (токени, повідомлення_про_помилку)."""
    tokens: List[str] = []
    for part in re.split(r"[\r\n,;\s]+", (text or "").strip()):
        t = norm_id(part)
        if not t:
            continue
        if len(t) < ID_MIN:
            return [], "Для пошуку потрібно щонайменше {} останніх символів: {}".format(
                ID_MIN, part
            )
        tokens.append(t)
    return tokens, None


def cell_has_id(s: str, tokens: Sequence[str]) -> bool:
    for p in _SPLIT_RE.split(s or ""):
        if not p or p.startswith("("):  # примітки в дужках не враховуються
            continue
        t = norm_id(p)
        if any(len(t) >= len(q) and t.endswith(q) for q in tokens):
            return True
    return False


def row_has_id(snap: Snapshot, r: int, tokens: Sequence[str]) -> bool:
    return any(
        cell_has_id(snap.text(r, c), tokens) for c in (COL_EIC, COL_METER)
    )


def filter_rows(
    snap: Snapshot,
    subst: str = ALL_SUBSTATIONS,
    text: str = "",
    id_tokens: Optional[Sequence[str]] = None,
) -> List[int]:
    needle = (text or "").strip().casefold()
    out: List[int] = []
    for r in snap.rows():
        if subst and subst != ALL_SUBSTATIONS and snap.subst_of(r) != subst:
            continue
        if needle and not any(needle in snap.text(r, c).casefold() for c in range(1, LAST_COL + 1)):
            continue
        if id_tokens and not row_has_id(snap, r, id_tokens):
            continue
        out.append(r)
    return out


def changed_fields(snap: Snapshot, r: int, new_values: Dict[int, str]) -> Dict[int, str]:
    """Поля, значення яких відрізняються від збережених; заблоковані пропускаються."""
    out: Dict[int, str] = {}
    for c, v in new_values.items():
        if snap.is_locked(r, c):
            continue
        if v != snap.text(r, c):
            out[c] = v
    return out


# ---------------------------------------------------------------- навігація
class Navigator:
    """Поточна вибірка рядків (фільтри) і позиція в ній — без Tk."""

    def __init__(self, snap: Snapshot) -> None:
        self.snap = snap
        self.subst = ALL_SUBSTATIONS
        self.text = ""
        self.id_tokens: Optional[List[str]] = None
        self.rows: List[int] = []
        self.idx = 0
        self.refresh()

    @property
    def current(self) -> Optional[int]:
        if 0 <= self.idx < len(self.rows):
            return self.rows[self.idx]
        return None

    def set_snapshot(self, snap: Snapshot) -> None:
        keep = self.current
        self.snap = snap
        self.refresh(keep)

    def refresh(self, keep_row: Optional[int] = None) -> None:
        if keep_row is None:
            keep_row = self.current
        self.rows = filter_rows(self.snap, self.subst, self.text, self.id_tokens)
        self.idx = 0
        if keep_row in self.rows:
            self.idx = self.rows.index(keep_row)

    def reset_filters(self) -> None:
        self.subst = ALL_SUBSTATIONS
        self.text = ""
        self.id_tokens = None
        self.refresh()

    def move(self, where: str) -> None:
        if not self.rows:
            return
        if where == "first":
            self.idx = 0
        elif where == "last":
            self.idx = len(self.rows) - 1
        elif where == "prev":
            self.idx = max(0, self.idx - 1)
        elif where == "next":
            self.idx = min(len(self.rows) - 1, self.idx + 1)

    def goto(self, row: int) -> bool:
        """Показати рядок листа; якщо він поза вибіркою — скинути фільтри."""
        if row not in self.rows:
            self.reset_filters()
        if row in self.rows:
            self.idx = self.rows.index(row)
            return True
        return False

    def find_connection(self, subst: str, name: str) -> Optional[int]:
        for r in self.snap.rows():
            if self.snap.subst_of(r) == subst and self.snap.text(r, COL_CONN).strip() == name:
                return r
        return None


def block_text(block: Block, texts: Sequence[str]) -> str:
    """Значення блоку для друку: «2022-03-14» → «14.03.2022», «III / 2027» тощо."""
    parts = [(t or "").strip() for t in texts]
    if block.sep == ".":
        parts = [p.zfill(2) if (i < 2 and p.isdigit()) else p for i, p in enumerate(parts)]
        return ".".join(parts) if any(parts) else ""
    if block.sep == "#":
        return " / ".join(p or "–" for p in parts) if any(parts) else ""
    if block.sep:
        return block.sep.join(p for p in parts if p)
    return parts[0] if parts else ""


# ---------------------------------------------------------------- друк карток
Card = List[Tuple[str, List[Tuple[str, str]]]]  # [(група, [(підпис, значення)])]


def build_card(
    snap: Snapshot,
    layout: Sequence[Group],
    r: int,
    override: Optional[Dict[int, str]] = None,
) -> Card:
    """Картка запису для друку; `override` — значення з полів форми (незбережені)."""
    over = override or {}
    card: Card = []
    for g in layout:
        items: List[Tuple[str, str]] = []
        for b in g.blocks:
            texts = [over.get(c, snap.text(r, c)) for c in b.cols]
            items.append((b.label, block_text(b, texts)))
        card.append((g.name, items))
    return card


_BAD_FILE_CHARS = re.compile(r'[\/:*?"<>|\r\n\t]+')


def safe_file_part(s: str) -> str:
    s = re.sub(r"\s+", " ", _BAD_FILE_CHARS.sub("-", s or ""))
    return re.sub(r"\s*-\s*", "-", s).strip(" .-")


def card_pdf_name(subst: str, conn: str, when: datetime) -> str:
    """«Дані по лічильнику - підстанція - приєднання - дд.мм.рррр гг-хх.pdf»."""
    parts = ["Дані по лічильнику", safe_file_part(subst), safe_file_part(conn)]
    stamp = when.strftime("%d.%m.%Y %H-%M")
    return " - ".join([p for p in parts if p] + [stamp]) + ".pdf"


# ---------------------------------------------------------------- журнал актів
SH_ACTS = "Акти"
ACT_COL_EXECUTOR = 56
ACT_COL_NUMBER = 57
ACT_COLS_DATE = (58, 59, 60)
ACT_COLS = (ACT_COL_EXECUTOR, ACT_COL_NUMBER) + ACT_COLS_DATE
ACT_HEADERS = [
    "Дата Час",
    "Користувач",
    "Подія",
    "Номер акту",
    "Дата акту",
    "Код виконавця",
    "Підстанція",
    "Приєднання",
    "EIC",
    "№ лічильника",
    "Вид роботи",
    "Вид перевірки",
    "Скан (файл)",
    "Рядок бази",
    "Примітка",
]
EVENT_RECORD_EDIT = "Правка запису"
EVENT_RENAME = "Скан перейменовано"


@dataclass
class ActEntry:
    """Один рядок журналу «Акти»: подія по акту, а не окрема зміна поля."""

    when: datetime
    user: str
    event: str
    act_no: str = ""
    act_date: str = ""
    executor: str = ""
    subst: str = ""
    conn: str = ""
    eic: str = ""
    meter: str = ""
    work_type: str = ""
    check_type: str = ""
    scan: str = ""
    sheet_row: Optional[int] = None
    note: str = ""

    def as_row(self) -> List[object]:
        return [
            self.when, self.user, self.event, self.act_no, self.act_date,
            self.executor, self.subst, self.conn, self.eic, self.meter,
            self.work_type, self.check_type, self.scan,
            self.sheet_row if self.sheet_row is not None else "", self.note,
        ]


def act_fields_changed(changes: Dict[int, str]) -> bool:
    return any(c in changes for c in ACT_COLS)


def act_date_text(day: str, month: str, year: str) -> str:
    parts = [(day or "").strip(), (month or "").strip(), (year or "").strip()]
    if not any(parts):
        return ""
    return ".".join(p.zfill(2) if (i < 2 and p.isdigit()) else p for i, p in enumerate(parts))


def act_entry_for_record(
    snap: Snapshot,
    r: int,
    changes: Optional[Dict[int, str]] = None,
    *,
    user: str = "",
    when: Optional[datetime] = None,
    event: str = EVENT_RECORD_EDIT,
    scan: str = "",
    work_type: str = "",
    check_type: str = "",
    note: str = "",
) -> ActEntry:
    """Рядок журналу актів для запису r; `changes` — значення, які щойно записуються."""
    over = changes or {}

    def txt(c: int) -> str:
        return (over[c] if c in over else snap.text(r, c)).strip()

    return ActEntry(
        when=when or datetime.now(),
        user=user,
        event=event,
        act_no=txt(ACT_COL_NUMBER),
        act_date=act_date_text(*(txt(c) for c in ACT_COLS_DATE)),
        executor=txt(ACT_COL_EXECUTOR),
        subst=snap.subst_of(r),
        conn=txt(COL_CONN),
        eic=txt(COL_EIC),
        meter=txt(COL_METER),
        work_type=work_type,
        check_type=check_type,
        scan=scan,
        sheet_row=r,
        note=note,
    )
