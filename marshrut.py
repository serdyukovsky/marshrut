# -*- coding: utf-8 -*-
"""
Маршрутный лист из МоегоСклада.

Берёт отгрузки за дату через API МоегоСклада и собирает готовый к печати Excel:
  - отгрузки крупным шрифтом, товар на строке клиента;
  - внизу посередине: наличка (отгрузки от организации «ГК Папирус Алтай») и итог;
  - внизу справа: сводка для производства (с пометкой «Обрез»).
Плюс печать: все реализации за день одним PDF, а где в комментарии есть
слово «счет» — к реализации добавляется счёт покупателю.

Настройки: config.ini рядом со скриптом.
Запуск: python marshrut.py [ДД.ММ.ГГГГ]   (без даты — сегодня)
"""
import configparser
import os
import sys
import time
from collections import OrderedDict
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

VERSION = "27.10 · 12:30"
BUILD = "2026.10.27.1230"  # номер для автообновления — тот же, что в version.txt на GitHub
UPDATE_BASE = "https://raw.githubusercontent.com/serdyukovsky/marshrut/main/"
BASE = "https://api.moysklad.ru/api/remap/1.2"
# папка программы: рядом с .exe (если собрано) или со скриптом
HERE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, "config.ini")
TOKEN_PLACEHOLDER = "ВСТАВЬ_ТОКЕН_СЮДА"
CONFIG_TEMPLATE = """[moysklad]
; Токен: МойСклад -> Настройки -> Токены (нужны права администратора).
; Никому не пересылай этот файл с токеном.
token = ВСТАВЬ_ТОКЕН_СЮДА

[marshrut]
; Организация, по которой идёт наличка (кавычки и регистр не важны)
cash_org = ГК Папирус Алтай

; Пометки справа от товара берутся из комментария к отгрузке.
; Сырьё — слова через запятую:
raw_words = Беларусь, Заринск, Красноярск
; «Обрез» ставится, если в комментарии есть это слово (обрез, обрезаем…)
obrez_word = обрез
; …и всегда для товаров, в названии которых есть:
obrez_products = 16 см

; Куда сохранять листы (пусто = рядом со скриптом)
out_dir =

[print]
; Шаблон реализации — название как в МоёмСкладе
demand_template = УПД 2026
demand_copies = 2

; Отгрузки по наличке (организация из cash_org) печатаются этим шаблоном
cash_template = Расходная накладная
cash_copies = 1

; ТТН-1: для этих организаций или контрагентов (через запятую) галочка «ТТН» стоит сразу
ttn_for = ТОРГСЕРВИС 22
ttn_template = ТТН-1 (Для торг)
ttn_copies = 2

; Счёт печатается, если в комментарии к отгрузке есть это слово (регистр и ё не важны)
invoice_word = счет
invoice_template = Счет покупателю
invoice_copies = 1

; open — открыть общий PDF (печать одной кнопкой Ctrl+P)
; print — сразу отправить на принтер по умолчанию
print_mode = open
"""


class Oops(Exception):
    """Понятная пользователю ошибка."""


# ---------------------------------------------------------------- настройки
def load_config():
    cfg = configparser.ConfigParser()
    path = CONFIG_PATH
    if not os.path.exists(path):  # первый запуск — создаём настройки по умолчанию
        with open(path, "w", encoding="utf-8") as f:
            f.write(CONFIG_TEMPLATE)
    cfg.read(path, encoding="utf-8")

    def lst(key):
        raw = cfg.get("marshrut", key, fallback="")
        return [x.strip().lower() for x in raw.split(",") if x.strip()]

    return {
        "token": cfg.get("moysklad", "token", fallback="").strip(),
        "cash_org": norm_org(cfg.get("marshrut", "cash_org", fallback="ГК Папирус Алтай")),
        "obrez": lst("obrez_products"),
        "obrez_word": yo(cfg.get("marshrut", "obrez_word", fallback="обрез")),
        "raw_words": [x.strip() for x in cfg.get("marshrut", "raw_words",
                      fallback="Беларусь, Заринск, Красноярск").split(",") if x.strip()],
        "out_dir": cfg.get("marshrut", "out_dir", fallback="").strip() or HERE,
        "demand_template": cfg.get("print", "demand_template", fallback="УПД 2026").strip(),
        "demand_copies": cfg.getint("print", "demand_copies", fallback=2),
        "cash_template": cfg.get("print", "cash_template", fallback="Расходная накладная").strip(),
        "cash_copies": cfg.getint("print", "cash_copies", fallback=1),
        "ttn_names": [norm_org(x) for x in cfg.get("print", "ttn_for", fallback="ТОРГСЕРВИС 22").split(",") if x.strip()],
        "ttn_template": cfg.get("print", "ttn_template", fallback="ТТН-1 (Для торг)").strip(),
        "ttn_copies": cfg.getint("print", "ttn_copies", fallback=2),
        "invoice_template": cfg.get("print", "invoice_template", fallback="Счет покупателю").strip(),
        "invoice_copies": cfg.getint("print", "invoice_copies", fallback=1),
        "invoice_word": yo(cfg.get("print", "invoice_word", fallback="счет")),
        "print_mode": cfg.get("print", "print_mode", fallback="open").strip().lower(),
        "depot_address": cfg.get("route", "depot_address", fallback="проспект Космонавтов 34Г, Барнаул").strip(),
        "depot_coords": cfg.get("route", "depot_coords", fallback="").strip(),
        "route_start": cfg.get("route", "start", fallback="09:00").strip(),
        "route_service": cfg.getint("route", "minutes_per_stop", fallback=10),
        "route_back": cfg.get("route", "return_to_depot", fallback="нет").strip().lower() in ("да", "yes", "1"),
        "fit_one_page": cfg.get("print", "fit_one_page", fallback="да").strip().lower() in ("да", "yes", "1", "true"),
    }


def has_token():
    try:
        t = load_config()["token"]
    except Exception:
        return False
    return bool(t) and TOKEN_PLACEHOLDER not in t.upper()


def check_token(token):
    """Проверяет токен в МоёмСкладе. Возвращает имя сотрудника."""
    try:
        r = session(token).get(BASE + "/context/employee", timeout=30)
    except Exception as e:
        raise Oops("Нет связи с МоимСкладом: %s" % e)
    if r.status_code in (401, 403):
        raise Oops("МойСклад не принял этот токен. Проверь, что скопирован целиком.")
    if not r.ok:
        raise Oops("МойСклад ответил ошибкой %s" % r.status_code)
    return r.json().get("name") or r.json().get("fullName") or "сотрудник"


def save_token(token):
    import re
    load_config()  # создаст файл, если его нет
    with open(CONFIG_PATH, encoding="utf-8") as f:
        text = f.read()
    text = re.sub(r"(?m)^token\s*=.*$", lambda m: "token = " + token.strip(), text, count=1)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write(text)


def save_setting(section, key, value):
    """Меняет одну настройку в config.ini, не трогая остальное и комментарии."""
    import re
    load_config()
    with open(CONFIG_PATH, encoding="utf-8") as f:
        text = f.read()
    line = "%s = %s" % (key, value)
    m = re.search(r"(?ms)^\[%s\]\s*$(.*?)(?=^\[|\Z)" % re.escape(section), text)
    if m:
        body = m.group(1)
        if re.search(r"(?m)^%s\s*=" % re.escape(key), body):
            body = re.sub(r"(?m)^%s\s*=.*$" % re.escape(key), lambda _: line, body, count=1)
        else:
            body = body.rstrip("\n") + "\n" + line + "\n\n"
        text = text[:m.start(1)] + body + text[m.end(1):]
    else:
        text = text.rstrip("\n") + "\n\n[%s]\n%s\n" % (section, line)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write(text)


def get_setting(section, key, default=""):
    cfg = configparser.ConfigParser()
    load_config()
    cfg.read(CONFIG_PATH, encoding="utf-8")
    return cfg.get(section, key, fallback=default)


def check_update():
    """Есть новая версия на GitHub — скачивает и подменяет marshrut.py. → номер новой версии или ""."""
    import re
    import shutil
    if get_setting("update", "auto", "yes").strip().lower() in ("no", "нет", "0", "off", "false"):
        return ""
    if os.path.exists(os.path.join(HERE, ".git")):  # папка разработки — не трогаем
        return ""
    try:
        import requests
        r = requests.get(UPDATE_BASE + "version.txt", timeout=4, headers={"Cache-Control": "no-cache"})
        r.raise_for_status()
        latest = r.text.strip()
        if not re.fullmatch(r"\d{4}\.\d{2}\.\d{2}\.\d{4}", latest) or latest <= BUILD:
            return ""
        r = requests.get(UPDATE_BASE + "marshrut.py", timeout=30, headers={"Cache-Control": "no-cache"})
        r.raise_for_status()
        code = r.content
        text = code.decode("utf-8")
        compile(text, "marshrut.py", "exec")  # файл целый и без ошибок
        if 'BUILD = "%s"' % latest not in text:
            return ""
        target = os.path.join(HERE, "marshrut.py")
        if os.path.exists(target):
            shutil.copy2(target, target + ".old")  # на всякий случай — предыдущая версия
        with open(target + ".new", "wb") as f:
            f.write(code)
        os.replace(target + ".new", target)
        return latest
    except Exception:
        log_error("проверка обновлений")
        return ""


def restart():
    """Перезапуск после обновления."""
    import subprocess
    if getattr(sys, "frozen", False):
        cmd = [sys.executable]
    else:
        cmd = [sys.executable, os.path.join(HERE, "marshrut.py")]
    subprocess.Popen(cmd, cwd=HERE)
    sys.exit(0)


def make_shortcut():
    """Ярлык на рабочем столе (Windows)."""
    if os.name != "nt":
        return False
    import subprocess
    target = sys.executable if getattr(sys, "frozen", False) else os.path.join(HERE, "Маршрутный лист.bat")
    ps = ("$d=[Environment]::GetFolderPath('Desktop');"
          "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\\Маршрутный лист.lnk');"
          "$s.TargetPath='%s';$s.WorkingDirectory='%s';$s.Save()" % (target.replace("'", "''"), HERE.replace("'", "''")))
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=20)
        return True
    except Exception:
        return False


def yo(t):
    return (t or "").lower().replace("ё", "е").strip()


def norm_org(t):
    """Название организации без кавычек, регистра и лишних пробелов."""
    import re
    return " ".join(re.sub(r"[^\w ]+", " ", yo(t)).split())


# ---------------------------------------------------------------- МойСклад
def session(token):
    import requests

    s = requests.Session()
    s.headers.update({
        "Authorization": "Bearer " + token,
        "Accept-Encoding": "gzip",
        "Accept": "application/json;charset=utf-8",
    })
    return s


def fetch_demands(token, day):
    s = session(token)
    d0 = day.strftime("%Y-%m-%d 00:00:00")
    d1 = day.strftime("%Y-%m-%d 23:59:59")
    rows, offset = [], 0
    while True:
        r = s.get(BASE + "/entity/demand", params={
            "filter": "moment>=%s;moment<=%s" % (d0, d1),
            "expand": "agent,organization,positions.assortment,invoicesOut,consignee",
            "limit": 100,
            "offset": offset,
            "order": "moment,desc",
        }, timeout=60)
        if r.status_code == 401:
            raise Oops("МойСклад не принял токен (401). Проверь token в config.ini.")
        if r.status_code == 403:
            raise Oops("У токена нет прав на чтение отгрузок (403).")
        r.raise_for_status()
        data = r.json()
        batch = data.get("rows", [])
        rows.extend(batch)
        if len(batch) < 100:
            break
        offset += 100
    load_uoms(s, rows)
    return rows


UOM = {}  # href единицы измерения → «шт», «упак» …


def load_uoms(s, demands):
    """Названия единиц измерения товаров и их фасовок (немного запросов, запоминаются)."""
    hrefs = set()
    for d in demands:
        for p in _positions(d):
            a = p.get("assortment") or {}
            for u in [a.get("uom")] + [pk.get("uom") for pk in a.get("packs") or []]:
                h = ((u or {}).get("meta") or {}).get("href")
                if h and h not in UOM:
                    hrefs.add(h)
    for h in hrefs:
        try:
            r = s.get(h, timeout=30)
            if r.ok:
                UOM[h] = (r.json().get("name") or "").strip()
        except Exception:
            pass


def uom_name(u, default):
    name = UOM.get(((u or {}).get("meta") or {}).get("href"), "")
    name = re.sub(r"^упак(?:овка)?\.?$", "уп", name, flags=re.I)
    return name or default


def pack_info(p):
    """Сколько штук в упаковке у позиции: выбранная в документе фасовка, иначе основная фасовка товара.
    → (штук в упаковке или None, «уп», «шт»)."""
    a = p.get("assortment") or {}
    packs = [pk for pk in a.get("packs") or [] if float(pk.get("quantity") or 0) > 1]
    chosen = None
    pack = p.get("pack")
    if pack:
        chosen = next((pk for pk in a.get("packs") or [] if pk.get("id") == pack.get("id")), pack)
    if not chosen and packs:
        named = [pk for pk in packs if re.match(r"уп", uom_name(pk.get("uom"), ""), re.I)]
        chosen = (named or sorted(packs, key=lambda pk: float(pk["quantity"])))[0]
    base = uom_name(a.get("uom"), "шт")
    if chosen and float(chosen.get("quantity") or 0) > 1:
        return float(chosen["quantity"]), uom_name(chosen.get("uom"), "уп"), base
    return None, "уп", base


def num(x):
    return int(x) if abs(x - round(x)) < 1e-9 else round(x, 2)


def qty_text(base_qty, per, pack_unit="уп", base_unit="шт"):
    """«2 уп (24 шт)», «2 уп + 6 шт (30 шт)», «6 шт»; без фасовки — «24 шт»."""
    if not per:
        return "%s %s" % (num(base_qty), base_unit)
    whole = int(base_qty // per + 1e-9)
    rest = num(base_qty - whole * per)
    if whole and not rest:
        return "%s %s (%s %s)" % (whole, pack_unit, num(base_qty), base_unit)
    if whole:
        return "%s %s + %s %s (%s %s)" % (whole, pack_unit, rest, base_unit, num(base_qty), base_unit)
    return "%s %s" % (rest, base_unit)


def _positions(demand):
    pos = demand.get("positions", {})
    if isinstance(pos, dict):
        return pos.get("rows", [])
    return pos or []


def packs_count(p):
    """Сколько упаковок (коробок) в позиции."""
    qty = float(p.get("quantity") or 0)
    pack = p.get("pack")
    if pack:
        per = pack.get("quantity")
        if not per:  # ищем упаковку в карточке товара
            for pk in (p.get("assortment") or {}).get("packs", []) or []:
                if pk.get("id") == pack.get("id"):
                    per = pk.get("quantity")
                    break
        if per:
            qty = qty / float(per)
    return int(qty) if abs(qty - round(qty)) < 1e-9 else round(qty, 2)


def attr_value(obj, *words):
    """Значение доп. поля, в названии которого есть все слова (без учёта регистра)."""
    for a in obj.get("attributes") or []:
        name = yo(a.get("name"))
        if all(w in name for w in words):
            v = a.get("value")
            if isinstance(v, dict):  # справочник
                v = v.get("name", "")
            if v not in (None, "", False):
                return str(v).strip()
    return ""


def first(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return ""


def normalize(demand, cfg):
    agent = demand.get("agent") or {}
    org = demand.get("organization") or {}
    agent_name = agent.get("name", "")
    # адрес — как в отчёте МоегоСклада: грузополучатель (если указан, иначе покупатель) →
    # его фактический адрес, а если его нет — юридический
    consignee = demand.get("consignee") or {}
    who = consignee if consignee.get("name") else agent
    address = first(who.get("actualAddress"), who.get("legalAddress"))
    address = address.replace("\u200b", "").strip()
    # комментарий к доставке — доп. поле отгрузки «Комментарий к доставке»
    delivery_comment = attr_value(demand, "коммент", "достав")
    delivery_comment = delivery_comment.replace("\u200b", "").strip()
    goods, key = [], []
    for p in _positions(demand):
        a = p.get("assortment") or {}
        per, pu, bu = pack_info(p)
        qty = float(p.get("quantity") or 0)
        goods.append({"name": a.get("name", "?"), "packs": packs_count(p), "qty": qty, "per": per,
                      "pu": pu, "bu": bu, "text": qty_text(qty, per, pu, bu)})
        key.append((href_id(a), round(float(p.get("quantity") or 0), 3)))
    invoices = [{"id": href_id(i), "name": i.get("name", "")} for i in demand.get("invoicesOut") or []]
    note = (demand.get("description") or "").strip()
    text = yo(note + " " + delivery_comment)
    raw = [w for w in cfg["raw_words"] if yo(w) in text]  # сырьё: Беларусь, Заринск…
    cut = bool(cfg["obrez_word"]) and cfg["obrez_word"] in text
    towels = [g for g in goods if "полотенц" in g["name"].lower()]
    for g in goods:
        marks = list(raw)
        if obrez(g["name"], cfg) or (cut and (g in towels or not towels)):
            marks.append("Обрез")
        g["marks"] = ", ".join(marks)
    return {
        "id": demand.get("id"),
        "meta": demand.get("meta"),
        "agent_meta": agent.get("meta"),
        "moment_full": demand.get("moment", "")[:19],
        "pos_key": sorted(key),
        "invoices": invoices,
        "number": demand.get("name", ""),
        "moment": demand.get("moment", "")[:16],
        "comment": delivery_comment,
        "agent": agent_name,
        "address": address,
        "note": note,
        "invoice": bool(cfg["invoice_word"]) and cfg["invoice_word"] in yo(note + " " + delivery_comment),
        "sum": (demand.get("sum") or 0) / 100.0,
        "org": org.get("name", ""),
        "cash": bool(cfg["cash_org"]) and cfg["cash_org"] in norm_org(org.get("name", "")),
        # ТТН по умолчанию — если организация или контрагент из списка ttn_for
        "ttn": any(n and (n in norm_org(org.get("name", "")) or n in norm_org(agent_name)) for n in cfg["ttn_names"]),
        "goods": goods,
        "hours": " \n ".join(x for x in [(agent.get("description") or ""), (consignee.get("description") or "")] if x),
    }


def href_id(obj):
    if obj.get("id"):
        return obj["id"]
    return (obj.get("meta") or {}).get("href", "").rstrip("/").split("/")[-1].split("?")[0]


def obrez(name, cfg):
    n = name.lower()
    return any(k in n for k in cfg["obrez"])


# ---------------------------------------------------------------- Excel
FONT = Font(name="Arial", size=12)
BOLD = Font(name="Arial", size=12, bold=True)
TITLE = Font(name="Arial", size=16, bold=True)
thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")

COLS = {"A": 5, "B": 11, "C": 13, "D": 21, "E": 28, "F": 28, "G": 49, "H": 19, "I": 13}


QTY_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def qty_cell(ws, ref, text):
    """Количество: упаковки обычным шрифтом, «(24 шт)» — мелко и серым, чтобы упаковки были главнее."""
    c = cell(ws, ref, text, align=QTY_ALIGN)
    if not isinstance(text, str) or " (" not in text:
        return c
    main, small = text.split(" (", 1)
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        c.value = CellRichText(TextBlock(InlineFont(rFont="Arial", sz=12), main + " "),
                               TextBlock(InlineFont(rFont="Arial", sz=9, color="666666"), "(" + small))
    except Exception:  # старый openpyxl — оставляем обычным текстом
        pass
    return c


def cell(ws, ref, value, font=FONT, border=True, align=WRAP, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = font
    c.alignment = align
    if border:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


def build_xlsx(orders, day, cfg, path, title=None, route_url=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Маршрутный лист"
    for col, w in COLS.items():
        ws.column_dimensions[col].width = w

    cell(ws, "B2", title or "Маршрутный лист  %s" % day.strftime("%d.%m.%Y"), TITLE, border=False, align=Alignment())
    if route_url:  # ссылка на маршрут в Яндекс Картах
        c = cell(ws, "G2", "Маршрут в Яндекс Картах", Font(name="Arial", size=12, color="1F6FEB", underline="single"),
                 border=False, align=Alignment(vertical="center"))
        c.hyperlink = route_url
    cell(ws, "A4", "№", BOLD, align=CENTER)
    heads = ["Номер", "Дата", "Комментарий к доставке", "Контрагент", "Адрес", "Товар", "Количество"]
    for i, h in enumerate(heads):
        cell(ws, "%s4" % "BCDEFGH"[i], h, BOLD)

    r = 5
    for n, o in enumerate(orders, 1):
        goods = o["goods"] or [{"name": "", "packs": ""}]
        dt = datetime.strptime(o["moment"][:10], "%Y-%m-%d") if o["moment"] else ""
        cell(ws, "A%d" % r, n, BOLD, align=CENTER)
        vals = [o["number"], dt, o["comment"], o["agent"], o["address"]]
        for i, v in enumerate(vals):
            cell(ws, "%s%d" % ("BCDEF"[i], r), v, fmt="DD.MM.YYYY" if i == 1 else None)
        for j, g in enumerate(goods):
            if j:
                for col in "ABCDEF":
                    cell(ws, "%s%d" % (col, r), None)
            cell(ws, "G%d" % r, g["name"])
            qty_cell(ws, "H%d" % r, g.get("text", g["packs"]))
            if g.get("marks"):
                cell(ws, "I%d" % r, g["marks"], BOLD, border=False, align=CENTER)
            r += 1
        r += 1  # пустая строка между клиентами

    # --- наличка (посередине) и производство (справа)
    r += 1
    cash = [o for o in orders if o["cash"]]
    rc = r
    for o in cash:
        cell(ws, "D%d" % rc, o["agent"])
        cell(ws, "E%d" % rc, o["sum"], align=CENTER, fmt="#,##0.00")
        rc += 1
    cell(ws, "D%d" % rc, "Итого:", BOLD, border=False)
    cell(ws, "E%d" % rc, sum(o["sum"] for o in cash), BOLD, align=CENTER, fmt="#,##0.00")

    rp = r
    for name, qty, marks in production_rows(orders):
        cell(ws, "G%d" % rp, name)
        qty_cell(ws, "H%d" % rp, qty)
        if marks:
            cell(ws, "I%d" % rp, marks, BOLD, border=False, align=CENTER)
        rp += 1

    ps = ws.page_setup
    ps.orientation = "landscape"
    ps.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ps.fitToWidth, ps.fitToHeight = 1, 0
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    wb.save(path)


# ---------------------------------------------------------------- печать
class RetrySession:
    """Обёртка над сессией: при 429 (лимит запросов) ждёт и повторяет."""

    def __init__(self, token):
        self.s = session(token)

    def request(self, method, url, **kw):
        for attempt in range(8):
            r = self.s.request(method, url, **kw)
            if r.status_code != 429:
                return r
            wait = r.headers.get("X-Lognex-Retry-TimeInterval")
            time.sleep(int(wait) / 1000.0 if wait and wait.isdigit() else 1.0 + attempt * 0.5)
        return r

    def get(self, url, **kw):
        return self.request("GET", url, **kw)

    def post(self, url, **kw):
        return self.request("POST", url, **kw)

    def put(self, url, **kw):
        return self.request("PUT", url, **kw)


PARALLEL = 4  # МойСклад разрешает не больше 5 одновременных запросов от пользователя


class Api:
    """Всё общение с МоимСкладом."""

    def __init__(self, cfg):
        import threading
        need_token(cfg)
        self.cfg = cfg
        self._local = threading.local()
        self._tpl = {}

    @property
    def s(self):  # своя сессия на каждый поток
        if not hasattr(self._local, "s"):
            self._local.s = RetrySession(self.cfg["token"])
        return self._local.s

    def _check(self, r):
        if r.status_code == 401:
            raise Oops("МойСклад не принял токен (401). Проверь token в config.ini.")
        if r.status_code == 403:
            raise Oops("У токена не хватает прав (403).")
        if r.status_code >= 400:
            raise Oops("МойСклад вернул ошибку %s: %s" % (r.status_code, r.text[:300]))
        return r

    def demands(self, day):
        try:
            return fetch_demands(self.cfg["token"], day)
        except Oops:
            raise
        except Exception as e:
            raise Oops("Не удалось связаться с МоимСкладом: %s" % e)

    def template(self, entity, name):
        """Ищет шаблон печати по названию среди своих и стандартных шаблонов."""
        if (entity, name) in self._tpl:
            return self._tpl[(entity, name)]
        found = []
        for kind in ("customtemplate", "embeddedtemplate"):
            r = self.s.get("%s/entity/%s/metadata/%s" % (BASE, entity, kind), timeout=60)
            if r.ok:
                found += r.json().get("rows", [])
        import re
        squash = lambda x: re.sub(r"[^0-9a-zа-я]+", "", yo(x))  # без пробелов, регистра и знаков
        want = squash(name)
        for exact in (True, False):
            for t in found:
                n = squash(t.get("name"))
                if (n == want) if exact else (want in n):
                    self._tpl[(entity, name)] = t["meta"]
                    return t["meta"]
        names = ", ".join(sorted({t.get("name", "") for t in found}))
        raise Oops("Не нашёл шаблон «%s». Есть такие: %s. Впиши нужное название в config.ini." % (name, names))

    def similar_invoice(self, o):
        """Счёт того же контрагента от той же даты с теми же позициями (даже если не связан)."""
        if not o["agent_meta"]:
            return None
        day = o["moment_full"][:10]
        r = self._check(self.s.get(BASE + "/entity/invoiceout", params={
            "filter": "agent=%s;moment>=%s 00:00:00;moment<=%s 23:59:59" % (o["agent_meta"]["href"], day, day),
            "expand": "positions", "limit": 100}, timeout=60))
        for inv in r.json().get("rows", []):
            key = sorted((href_id(p.get("assortment") or {}), round(float(p.get("quantity") or 0), 3))
                         for p in _positions(inv))
            if key == o["pos_key"]:
                return {"id": inv["id"], "name": inv.get("name", "")}
        return None

    def create_invoice(self, o):
        """Создаёт счёт на основании отгрузки (он сразу связан с ней)."""
        tpl = self._check(self.s.put(BASE + "/entity/invoiceout/new",
                                     json={"demands": [{"meta": o["meta"]}]}, timeout=60)).json()
        pos = tpl.get("positions")
        if isinstance(pos, dict):
            tpl["positions"] = pos.get("rows", [])
        tpl["moment"] = o["moment_full"] or tpl.get("moment")
        inv = self._check(self.s.post(BASE + "/entity/invoiceout", json=tpl, timeout=60)).json()
        return {"id": inv["id"], "name": inv.get("name", "")}

    def export(self, entity, doc_id, template_meta, count):
        """Печатная форма документа в PDF (count экземпляров)."""
        body = {"templates": [{"template": {"meta": template_meta}, "count": count}]}
        r = self.s.post("%s/entity/%s/%s/export/" % (BASE, entity, doc_id), json=body,
                        allow_redirects=False, timeout=120)
        for _ in range(60):  # 202 — форма ещё готовится
            if r.status_code != 202:
                break
            time.sleep(0.4)
            r = self.s.get(r.headers["Location"], allow_redirects=False, timeout=60)
        if r.status_code != 303 or "Location" not in r.headers:
            raise Oops("МойСклад не отдал печатную форму (код %s): %s" % (r.status_code, r.text[:200]))
        import requests
        f = requests.get(r.headers["Location"], timeout=120)  # временная ссылка, без токена
        f.raise_for_status()
        return f.content


def dump_fields(api, day):
    """Диагностика: все текстовые поля отгрузки, заказа покупателя и контрагента (без товаров и сумм)."""
    import json

    def flat(obj, skip=("meta", "positions", "rate", "files", "owner", "group", "state", "store",
                        "organizationAccount", "agentAccount", "accountId", "syncId", "sum", "vatSum",
                        "payedSum", "shippedSum", "invoicedSum", "reservedSum")):
        out = {}
        for k, v in (obj or {}).items():
            if k in skip or v in (None, "", [], {}):
                continue
            if k == "attributes":
                out["Доп. поля"] = {a.get("name"): (a["value"].get("name") if isinstance(a.get("value"), dict)
                                                     else a.get("value")) for a in v}
            elif isinstance(v, str):
                out[k] = v
            elif isinstance(v, dict) and k.endswith("Full"):
                out[k] = {kk: vv for kk, vv in v.items() if isinstance(vv, str)}
            elif isinstance(v, dict) and v.get("name"):
                out[k] = v["name"]
        return out

    s = api.s
    rows = []
    for d in api.demands(day):
        order = d.get("customerOrder") or {}
        if order.get("meta") and not order.get("name"):
            r = s.get(order["meta"]["href"], timeout=60)
            order = r.json() if r.ok else order
        agent = d.get("agent") or {}
        rows.append({"Отгрузка": flat(d), "Заказ покупателя": flat(order) if order else "нет",
                     "Контрагент": flat(agent),
                     "Грузополучатель": flat(d.get("consignee")) if d.get("consignee") else "нет"})
    path = os.path.join(HERE, "Диагностика %s.json" % day.strftime("%d.%m.%Y"))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    return path


def load_orders(api, cfg, day):
    orders = [normalize(d, cfg) for d in api.demands(day)]
    if not orders:
        raise Oops("За %s отгрузок не найдено." % day.strftime("%d.%m.%Y"))
    return orders


def plan_print(api, cfg, day, progress=lambda t: None):
    """Список документов на печать: по каждой отгрузке УПД и (если надо) счёт.
    Счёт: связанный со отгрузкой > найденный такой же > будет создан новый."""
    return annotate(api, sort_like_sheet(cfg, day, load_orders(api, cfg, day)), progress)


def annotate(api, orders, progress=lambda t: None, only=None):
    """Проставляет, что печатать, и ищет существующие счета. only — id отгрузок, отмеченных к печати."""
    for i, o in enumerate(orders, 1):
        on = only is None or o["id"] in only
        o["print_upd"] = on
        o["print_inv"] = o["invoice"] and on
        o["print_ttn"] = o["ttn"] and on
        if o["invoices"]:
            o["inv"] = dict(o["invoices"][0], kind="linked")
        elif o["invoice"]:
            progress("Проверяю счета %d из %d: %s" % (i, len(orders), o["agent"]))
            found = api.similar_invoice(o)
            o["inv"] = dict(found, kind="similar") if found else {"kind": "new"}
        else:
            o["inv"] = {"kind": "new"}
    return orders


def inv_label(o):
    inv = o["inv"]
    if inv["kind"] == "linked":
        return "уже есть № %s" % inv["name"]
    if inv["kind"] == "similar":
        return "такой же уже есть № %s" % inv["name"]
    return "создать новый"


def doc_for(o, cfg):
    """Какую накладную печатать: по наличке (ГК) — расходная, иначе УПД. -> (подпись, шаблон, копий)"""
    if o["cash"]:
        return "Расходная", cfg["cash_template"], cfg["cash_copies"]
    return "УПД", cfg["demand_template"], cfg["demand_copies"]


class Prefetcher:
    """Готовит печатные формы заранее, в фоне (пока человек смотрит список)."""

    def __init__(self, api):
        from concurrent.futures import ThreadPoolExecutor
        import threading
        self.api = api
        self.pool = ThreadPoolExecutor(max_workers=PARALLEL)
        self.futs = {}
        self.lock = threading.Lock()

    def _job(self, entity, doc_id, tpl_name, copies):
        tpl_entity = "invoiceout" if entity == "invoiceout" else "demand"
        return self.api.export(entity, doc_id, self.api.template(tpl_entity, tpl_name), copies)

    def get(self, entity, doc_id, tpl_name, copies):
        key = (entity, doc_id, tpl_name, copies)
        with self.lock:
            f = self.futs.get(key)
            if f is None or (f.done() and f.exception() is not None):  # ошибка — пробуем заново
                f = self.pool.submit(self._job, *key)
                self.futs[key] = f
            return f

    def warm(self, orders, cfg):
        """Запускает подготовку того, что отмечено по умолчанию (кроме ещё не созданных счетов)."""
        for o in orders:
            if o.get("print_upd"):
                _, name, copies = doc_for(o, cfg)
                self.get("demand", o["id"], name, copies)
            if o.get("print_ttn"):
                self.get("demand", o["id"], cfg["ttn_template"], cfg["ttn_copies"])
            if o.get("print_inv") and o["inv"]["kind"] != "new":
                self.get("invoiceout", o["inv"]["id"], cfg["invoice_template"], cfg["invoice_copies"])

    def progress(self):
        with self.lock:
            fs = list(self.futs.values())
        return sum(1 for f in fs if f.done()), len(fs)

    def close(self):
        self.pool.shutdown(wait=False, cancel_futures=True)


A4_PORTRAIT = (595.28, 841.89)
FIT_NOTES = []  # что пошло не так при ужимании — покажем в окне


def log_error(where):
    """Пишет полный текст ошибки в «ошибки.log» рядом с программой."""
    import traceback
    try:
        with open(os.path.join(HERE, "ошибки.log"), "a", encoding="utf-8") as f:
            f.write("\n[%s] %s\n%s" % (datetime.now().strftime("%d.%m.%Y %H:%M:%S"), where, traceback.format_exc()))
    except Exception:
        pass


def content_bbox(page, sc):
    """Границы не белого содержимого страницы в пикселях (без Pillow): (x0, y0, x1, y1) или None."""
    import re
    bm = page.render(scale=sc, grayscale=True)
    buf, w, h, st = bytes(bm.buffer), bm.width, bm.height, bm.stride
    dark = re.compile(rb"[\x00-\xf4]")
    top = bot = None
    left, right = w, -1
    for y in range(h):
        row = buf[y * st:y * st + w]
        m = dark.search(row)
        if not m:
            continue
        if top is None:
            top = y
        bot = y
        left = min(left, m.start())
        right = max(right, w - 1 - dark.search(row[::-1]).start())
    return None if top is None else (left, top, right + 1, bot + 1)


def fit_one_page(pdf_bytes, copies=1, min_scale=0.45, page_size=None, force=False):
    """Если один экземпляр документа занимает 2+ листа — обрезает пустые поля, ставит листы
    друг под другом и уменьшает, чтобы экземпляр поместился на один лист.
    page_size — размер итогового листа (например, A4 книжный); force — переложить даже одностраничный."""
    import io
    from pypdf import PdfReader, PdfWriter, PageObject, Transformation

    reader = PdfReader(io.BytesIO(pdf_bytes))
    total = len(reader.pages)
    copies = max(1, copies)
    if total % copies or (total <= copies and not force):
        return pdf_bytes  # и так по листу на экземпляр (или непонятная раскладка)
    per = total // copies

    # поворот страниц переносим в содержимое, чтобы координаты совпадали
    norm = PdfWriter()
    for pg in reader.pages:
        pg.transfer_rotation_to_content()
        norm.add_page(pg)
    nb = io.BytesIO()
    norm.write(nb)
    reader = PdfReader(io.BytesIO(nb.getvalue()))
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(nb.getvalue())
    except Exception as e:  # без библиотеки не обрезаем поля, но в один лист всё равно сводим
        doc = None
        log_error("pypdfium2")
        note = "не загрузилась библиотека pypdfium2 (%s) — поля не обрезаются" % e.__class__.__name__
        if note not in FIT_NOTES:
            FIT_NOTES.append(note)

    # границы содержимого каждой страницы (в пунктах PDF, от нижнего левого угла)
    boxes = []
    for i in range(total):
        if doc is None:
            mb = reader.pages[i].mediabox
            boxes.append((0, 0, float(mb.width), float(mb.height)))
            continue
        pg = doc[i]
        w, h = pg.get_size()
        sc = 0.75
        bbox = content_bbox(pg, sc)
        if not bbox:
            boxes.append(None)
            continue
        m = 6  # небольшой отступ
        x0, y0, x1, y1 = bbox
        boxes.append((max(0, x0 / sc - m), max(0, h - y1 / sc - m), min(w, x1 / sc + m), min(h, h - y0 / sc + m)))

    out = PdfWriter()
    for c in range(copies):
        idx = list(range(c * per, (c + 1) * per))
        parts = [(i, boxes[i]) for i in idx if boxes[i]]
        first = reader.pages[idx[0]]
        W, H = page_size or (float(first.mediabox.width), float(first.mediabox.height))
        if not parts:
            out.add_page(first)
            continue
        need_h = sum(b[3] - b[1] for _, b in parts)
        need_w = max(b[2] - b[0] for _, b in parts)
        margin = 18
        scale = min(1.0, (W - 2 * margin) / need_w, (H - 2 * margin) / need_h)
        if scale < min_scale:  # слишком мелко — оставляем как было
            for i in idx:
                out.add_page(reader.pages[i])
            continue
        page = PageObject.create_blank_page(width=W, height=H)
        y = H - margin
        x_left = (W - need_w * scale) / 2
        for i, (x0, y0, x1, y1) in parts:
            src = reader.pages[i]
            ph = (y1 - y0) * scale
            y -= ph
            t = Transformation().translate(-x0, -y0).scale(scale, scale).translate(x_left, y)
            page.merge_transformed_page(src, t)
        out.add_page(page)

    buf = io.BytesIO()
    out.write(buf)
    return buf.getvalue()


def do_print(api, cfg, day, orders, progress=lambda t: None, pf=None, cancel=None):
    """Печатает отмеченное одним PDF. Возвращает (путь, упд, счетов, создано счетов)."""
    from pypdf import PdfReader, PdfWriter
    import io

    chosen = [o for o in orders if o["print_upd"] or o["print_inv"] or o.get("print_ttn")]
    del FIT_NOTES[:]
    if not chosen:
        raise Oops("Ничего не отмечено для печати.")
    own_pf = pf is None
    pf = pf or Prefetcher(api)

    # 1) новые счета — по очереди, чтобы не создать дублей
    created = 0

    def stop_if_cancelled():
        if cancel is not None and cancel.is_set():
            msg = "Печать отменена."
            if created:
                msg += " Уже созданные счета (%d) остались в МоёмСкладе." % created
            raise Oops(msg)

    for o in chosen:
        stop_if_cancelled()
        if o["print_inv"] and o["inv"]["kind"] == "new":
            progress("Создаю счёт: %s" % o["agent"])
            o["inv"] = dict(api.create_invoice(o), kind="linked")
            created += 1

    # 2) печатные формы — параллельно (часть уже готова заранее), собираем в порядке листа
    from concurrent.futures import wait, FIRST_COMPLETED
    futs = {}
    n_copies = {}
    for i, o in enumerate(chosen):
        if o["print_upd"]:
            _, name, copies = doc_for(o, cfg)
            futs[pf.get("demand", o["id"], name, copies)] = (i, 0)
            n_copies[(i, 0)] = copies
        if o["print_inv"]:
            futs[pf.get("invoiceout", o["inv"]["id"], cfg["invoice_template"], cfg["invoice_copies"])] = (i, 2)
            n_copies[(i, 2)] = cfg["invoice_copies"]
        if o.get("print_ttn"):
            futs[pf.get("demand", o["id"], cfg["ttn_template"], cfg["ttn_copies"])] = (i, 1)
            n_copies[(i, 1)] = cfg["ttn_copies"]
    pdfs, done = {}, 0
    try:
        pending = set(futs)
        while pending:
            stop_if_cancelled()
            finished, pending = wait(pending, timeout=0.3, return_when=FIRST_COMPLETED)
            for f in finished:
                pdfs[futs[f]] = f.result()
                done += 1
            if finished:
                progress("Печатаю %d из %d документов…" % (done, len(futs)))
        stop_if_cancelled()
    finally:
        if own_pf:
            pf.close()

    out = PdfWriter()
    for key in sorted(pdfs):
        data = pdfs[key]
        try:
            if key[1] == 1:  # ТТН — всегда один книжный лист A4 на экземпляр
                data = fit_one_page(data, n_copies.get(key, 1), min_scale=0, page_size=A4_PORTRAIT, force=True)
            elif cfg["fit_one_page"]:  # документ на 2 листа — ужимаем в один
                data = fit_one_page(data, n_copies.get(key, 1))
        except Exception as e:
            log_error("ужимание документа")
            FIT_NOTES.append("не удалось ужать документ: %s" % e)
        for page in PdfReader(io.BytesIO(data)).pages:
            out.add_page(page)
    n_upd = sum(1 for o in chosen if o["print_upd"] and not o["cash"])
    n_cash = sum(1 for o in chosen if o["print_upd"] and o["cash"])
    n_inv = sum(1 for o in chosen if o["print_inv"])

    path = free_path(cfg, "Документы %s" % day.strftime("%d.%m.%Y"), day, ".pdf")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, "wb") as f:
            out.write(f)
    except PermissionError:
        raise Oops("Файл «%s» открыт — закрой его и попробуй снова." % os.path.basename(path))
    return (path, n_upd, n_inv, created, n_cash, sum(1 for o in chosen if o.get("print_ttn")),
            "; ".join(dict.fromkeys(FIT_NOTES)))


def send_to_print(path, mode):
    if os.name != "nt":
        import subprocess
        if mode == "print" and sys.platform == "darwin":
            try:
                subprocess.run(["lp", path], check=True, timeout=60)
                return
            except Exception:
                pass
        return open_file(path)
    if mode == "print":
        try:
            os.startfile(path, "print")
            return
        except OSError:
            pass
    os.startfile(path)


def need_token(cfg):
    if not cfg["token"] or "ВСТАВЬ" in cfg["token"].upper():
        raise Oops("Нет токена МоегоСклада — нажми «Настройки» вверху справа и вставь его.")




# ---------------------------------------------------------------- маршрут
# Бесплатно и без ключей: адрес → координаты через OpenStreetMap (Nominatim),
# время и расстояние по дорогам — через OSRM. Найденные точки запоминаются в «адреса.json».
GEO_PATH = os.path.join(HERE, "адреса.json")
UA = "PapirusAltai-Marshrut/1.0 (route sheet tool)"
BARNAUL_BOX = "83.45,53.50,84.05,53.20"  # левый-верхний, правый-нижний угол (долгота, широта)
ALTAI_BOX = (49.5, 77.5, 54.7, 90.0)  # Алтайский край + Республика Алтай: юг, запад, север, восток


def in_altai(ll):
    return bool(ll) and ALTAI_BOX[0] <= ll[0] <= ALTAI_BOX[2] and ALTAI_BOX[1] <= ll[1] <= ALTAI_BOX[3]


def in_barnaul(ll):
    return bool(ll) and 53.22 <= ll[0] <= 53.48 and 83.45 <= ll[1] <= 84.05
_last_nominatim = [0.0]


import re  # noqa: E402  (нужен в geo_load)


def geo_load():
    import json
    try:
        with open(GEO_PATH, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        return {}
    # чистим явные ошибки старых версий: точка вне края; село, «найденное» в Барнауле
    for k in list(cache):
        if k.startswith("hours|"):
            continue  # сохранённые часы работы
        v = cache[k] or {}
        ll = v.get("ll")
        if not ll:
            cache.pop(k)
            continue
        if not in_altai(ll):
            cache.pop(k)
            continue
        text = k[2:] if k.startswith("q:") else (v.get("found") or "")
        try:
            p = parse_address(text)
        except Exception:
            continue
        rural = bool(p["district"])  # «Ключевский р-н, село Ключи» точно не в Барнауле (Южный, Власиха — могут)
        if rural and in_barnaul(ll) and not re.search(r"барнаул", v.get("found_map", "") or "", re.I) \
                and k != "цех":
            cache.pop(k)
    return cache


def geo_save(cache):
    import json
    try:
        with open(GEO_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
    except OSError:
        pass


def geo_key(o):
    """По какому ключу помним точку: контрагент/грузополучатель + адрес."""
    return yo("%s | %s" % (o.get("agent", ""), " ".join((o.get("address") or "").split())))


def parse_coords(text):
    """«53.3481, 83.7757» (как копирует Яндекс/Google Карты) → (lat, lon) или None."""
    from urllib.parse import unquote
    text = unquote(text or "")  # ссылка из Яндекс Карт: ...ll=83.776%2C53.346...
    m = re.search(r"(-?\d{1,3}[.,]\d{2,})\s*[,; ]\s*(-?\d{1,3}[.,]\d{2,})", text)
    if not m:
        return None
    lat, lon = float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", "."))
    if 40 < lon < 120 and 40 < lat < 70:
        return lat, lon
    if 40 < lat < 120 and 40 < lon < 70:  # перепутаны местами
        return lon, lat
    return None


STREET_TYPES = [  # (как пишут, как понимает карта)
    (r"пр-кт|пр-т|просп\.?|проспект|пр\.", "проспект"), (r"ул\.?|улица", "улица"), (r"пер\.?|переулок", "переулок"),
    (r"б-р|бул\.?|бульвар", "бульвар"), (r"пл\.?|площадь", "площадь"), (r"ш\.?|шоссе", "шоссе"),
    (r"тракт", "тракт"), (r"проезд|пр-д", "проезд"), (r"мкр\.?|микрорайон", "микрорайон"), (r"наб\.?|набережная", "набережная"),
]
PLACE_TYPES = r"(?i:(?:город|г\.?|село|сел\.|с\.?|посёлок|поселок|пос\.?|пгт|рп|р\.п\.|п\.|деревня|дер\.?|станция|ст\.?)(?![а-яё]))"


def parse_address(addr):
    """Разбирает адрес на части: населённый пункт, район, улица, дом. Работает и с «Карла Маркса ул,95»."""
    import re
    a = " ".join((addr or "").replace("​", " ").replace("\n", " ").split())
    a = re.sub(r"(\+7|8)[\s\-()]*\d{3}[\s\-()]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}", " ", a)  # телефоны
    a = re.sub(r"\bтел\.?\s*[\d\-\s]{5,}", " ", a, flags=re.I)
    a = re.sub(r"\b\d{6}\b,?", " ", a)  # индекс
    a = re.sub(r"\b(?:ИНН|КПП|ОГРНИП|ОГРН)\b.*$", " ", a, flags=re.I)  # реквизиты в хвосте адреса
    region = "Алтайский край"
    if re.search(r"\bРесп(?:ублика|\.)?\s*Алтай\b|Горно-Алтайск", a, re.I):
        region = "Республика Алтай"
    a = re.sub(r"\bРоссия\b,?|\bАлтайский\s+край\b,?|\bАлтайский\s+кр\.?,?|\bРесп(?:ублика|\.)?\s*Алтай\b\.?,?",
               " ", a, flags=re.I)
    a = re.sub(r"([а-яё])([А-ЯЁ])", r"\1 \2", a)  # «центрПроспект» → «центр Проспект»
    a = re.sub(r"(\d)([а-яё]{3,})", r"\1 \2", a)  # «к7цокольный» → «к7 цокольный»
    a = re.sub(r"\b(?:цокольный|подвальный|\d+-?й)\s+этаж\b|\bэтаж\s*\d+|\b\d+\s*этаж\b", " ", a, flags=re.I)
    a = re.sub(r"\b(?:к|корп\.?|корпус|стр\.?|строение|лит\.?|литер)\s*\d+[а-я]?\b", " ", a, flags=re.I)
    a = re.sub(r",?\s*\b(?:пом|помещение|оф|офис|каб|кабинет|кв)\.?\s*\d+\S*", " ", a, flags=re.I)  # помещение/офис
    a = re.sub(r"\bд\.?\s*(?=\d)", "", a)  # «д. 126а» → «126а»
    a = " ".join(a.split()).strip(" ,.")
    res = {"clean": a, "place": None, "district": None, "street": None, "house": None, "region": region}
    m = re.search(r"([А-ЯЁ][а-яё\-]+(?:ий|ый|ой))\s+(?:р-н|район)", a)
    if m:
        res["district"] = m.group(1) + " район"
    for part in [x.strip(" .") for x in re.split(r",|(?<=[а-яё]{3})\.\s+", a)]:  # «… Алтай. г. Горно-Алтайск»
        m = re.match(r"^%s\s*([А-ЯЁ][А-Яа-яЁё\-]+(?:\s[А-ЯЁ][а-яё\-]+)?)$" % PLACE_TYPES, part)
        if m and not re.match(r"^(ул|пр|пер|пл|д)\b", part, re.I):
            res["place"] = m.group(1)
            continue
        m = re.match(r"^([А-ЯЁ][а-яё\-]+)\s+(?:с|село)\.?$", part)  # «Родино с»
        if m:
            res["place"] = m.group(1)
    types = "(?:%s)(?![а-яё])" % "|".join(t for t, _ in STREET_TYPES)  # «Проспект» ≠ «просп» + «ект»
    house = r"(?:д\.?\s*|дом\s*|зд\.?\s*)?(\d+[а-яА-Я]?(?:\s*/\s*\d+[а-яА-Я]?)?)"
    m = re.search(r"(?:^|[\s,])(%s)\s*([А-ЯЁа-яё0-9][А-Яа-яЁё0-9.\- ]*?)\s*,?\s*%s\b" % (types, house), a, re.I) \
        or None
    if m:
        kind, name, num = m.group(1), m.group(2), m.group(3)
    else:  # «Карла Маркса ул, 95», «Павловский тракт, 188»
        m = re.search(r"([А-ЯЁа-яё0-9][А-Яа-яЁё0-9.\- ]*?)\s+(%s)\s*,?\s*%s\b" % (types, house), a, re.I)
        kind, name, num = (m.group(2), m.group(1), m.group(3)) if m else (None, None, None)
    if m:
        full = next(f for t, f in STREET_TYPES if re.fullmatch(t, kind, re.I))
        name = re.sub(r"^(?:г\.?\s*)?Барнаул\.?\s*", "", name.strip(" ,."), flags=re.I)
        res["street"] = "%s %s" % (full, name) if full not in ("тракт",) else "%s %s" % (name, full)
        res["house"] = num.replace(" ", "")
    if not m:  # тип улицы не указан: «Автотранспортная, 49И», «Социалистический 4», «Дзержинского, 41»
        for g in re.finditer(r"(?:^|,\s*)([А-ЯЁ][а-яё\-]+(?:\s[А-ЯЁ][а-яё\-]+)?)\s*,?\s*(\d+[а-яА-Я]?(?:/\d+)?)\b", a):
            name, num = g.group(1), g.group(2)
            if re.search(r"(ая|яя)$", name):
                res["street"], res["house"] = "улица " + name, num
            elif re.search(r"(ий|ый|ой)$", name):
                res["street"], res["house"] = name + " проспект", num
            elif re.search(r"(ого|его|ова|ева|ёва|ина|ына|ых|их|ей|ов|ев)$", name):  # «Строителей 5/1»
                res["street"], res["house"] = "улица " + name, num
            else:
                continue
            break
    if not res["place"] and re.search(r"\bбарнаул", a, re.I):
        res["place"] = "Барнаул"
    return res


def address_queries(addr):
    """Варианты запроса к карте — от самого точного к приблизительному. → [(подпись, параметры)]"""
    import re
    p = parse_address(addr)
    out_of_town = bool(p["district"]) or bool(p["place"] and p["place"].lower() != "барнаул")
    place = p["place"] or ("" if out_of_town else "Барнаул")
    region = p.get("region") or "Алтайский край"
    out_of_town = out_of_town or region != "Алтайский край"
    tail = ", ".join(x for x in [place, p["district"], region] if x)
    qs = []
    if p["street"] and p["house"]:
        qs.append(("%s %s, %s" % (p["street"], p["house"], tail),
                   {"street": "%s %s" % (p["house"], p["street"]), "city": place, "county": p["district"] or "",
                    "state": region, "country": "Россия"}))
        qs.append(("%s %s, %s" % (p["street"], p["house"], tail), {"q": "%s %s, %s" % (p["street"], p["house"], tail)}))
    clean = p["clean"]
    if clean:
        q = clean if (place.lower()[:6] in clean.lower()) else "%s, %s" % (clean, tail)
        qs.append((q, {"q": q}))
        nb = " ".join(re.sub(r"\([^)]*\)?|\"[^\"]*\"", " ", q).split()).strip(" ,.")
        if nb != q:
            qs.append((nb, {"q": nb}))
    if p["street"]:  # без номера дома — хотя бы улица
        qs.append(("%s, %s (без дома)" % (p["street"], tail), {"q": "%s, %s" % (p["street"], tail)}))
    if out_of_town and place:  # в сёлах дома часто не отмечены — хотя бы центр села
        qs.append(("%s (центр)" % tail, {"q": tail}))
    seen, res = set(), []
    for label, params in qs:
        k = tuple(sorted((a, b) for a, b in params.items() if b))
        if k not in seen:
            seen.add(k)
            res.append((label, {a: b for a, b in params.items() if b}))
    return res


def address_variants(addr):
    """Подписи вариантов (для проверки)."""
    return [label for label, _ in address_queries(addr)]


def nominatim(params):
    """Один запрос к OpenStreetMap (не чаще раза в секунду — таковы правила сервиса)."""
    import requests
    wait = 1.1 - (time.time() - _last_nominatim[0])
    if wait > 0:
        time.sleep(wait)
    _last_nominatim[0] = time.time()
    box = "%s,%s,%s,%s" % (ALTAI_BOX[1], ALTAI_BOX[2], ALTAI_BOX[3], ALTAI_BOX[0])  # Алтайский край + Республика Алтай
    q = dict(params, format="jsonv2", limit=1, countrycodes="ru", viewbox=box, bounded=1, **{"accept-language": "ru"})
    r = requests.get("https://nominatim.openstreetmap.org/search", headers={"User-Agent": UA}, params=q, timeout=20)
    r.raise_for_status()
    rows = r.json()
    for row in rows:
        ll = (float(row["lat"]), float(row["lon"]))
        if in_altai(ll):
            return ll[0], ll[1], row.get("display_name", "")
    return None


def geocode(addr, cache, key=None, progress=lambda t: None):
    """Координаты адреса: из памяти или из OpenStreetMap. → (lat, lon) или None.
    Найденное запоминается; «не найдено» — нет (чтобы в следующий раз попробовать снова)."""
    import json
    if key and key in cache and cache[key].get("ll"):
        return tuple(cache[key]["ll"])
    direct = parse_coords(addr)
    if direct:
        return direct
    for label, params in address_queries(addr):
        ck = "q:" + json.dumps(params, ensure_ascii=False, sort_keys=True)
        if cache.get(ck, {}).get("ll"):
            return tuple(cache[ck]["ll"])
        progress("Ищу на карте: %s" % label[:70])
        try:
            got = nominatim(params)
        except Exception:
            log_error("поиск адреса")
            raise Oops("Не получилось связаться с картой OpenStreetMap. Проверь интернет и попробуй ещё раз.")
        if got:
            cache[ck] = {"ll": [got[0], got[1]], "found_map": got[2],
                         "approx": "(центр)" in label or "(без дома)" in label}
            return got[0], got[1]
    return None


def haversine_km(a, b):
    import math
    la1, lo1, la2, lo2 = map(math.radians, (a[0], a[1], b[0], b[1]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 6371 * 2 * math.asin(math.sqrt(h))


def road_matrix(points):
    """Минуты и км по дорогам между всеми точками (OSRM). Если сервис недоступен — по прямой ×1.4."""
    import requests
    n = len(points)
    try:
        coords = ";".join("%.6f,%.6f" % (p[1], p[0]) for p in points)
        r = requests.get("https://router.project-osrm.org/table/v1/driving/" + coords,
                         params={"annotations": "duration,distance"}, headers={"User-Agent": UA}, timeout=30)
        r.raise_for_status()
        j = r.json()
        if j.get("code") != "Ok":
            raise ValueError(j.get("code"))
        mins = [[(d or 0) / 60.0 for d in row] for row in j["durations"]]
        kms = [[(d or 0) / 1000.0 for d in row] for row in j["distances"]]
        return mins, kms, True
    except Exception:
        log_error("расчёт дорог")
        kms = [[haversine_km(points[i], points[j]) * 1.4 for j in range(n)] for i in range(n)]
        mins = [[k / 25.0 * 60 for k in row] for row in kms]  # ~25 км/ч по городу
        return mins, kms, False


WEEKDAYS = [("понедельн", "пн"), ("вторн", "вт"), ("сред", "ср"), ("четверг", "чт"), ("пятниц", "пт"),
            ("суббот", "сб"), ("воскрес", "вс")]


def time_window(text, day):
    """Окно доставки из комментариев: «до 16:00», «с 10-30 до 17-00», «по пятницам до 14-00».
    → (с, до) в минутах от полуночи; None — без ограничения."""
    import re
    hhmm = r"(\d{1,2})(?:\s*[:.\-]\s*(\d{2}))?"
    earliest, latest = None, None
    wd = day.weekday()
    for line in re.split(r"[\n;]+", yo(text or "")):
        # «касса до 13-00», «оплата до…», «бухгалтерия…», «обед с 13 до 14» — не про приём товара
        line = re.sub(r"\b(?:касс|оплат|бухгалтер|бух\b|обед|перерыв)[^,.;]*", " ", line)
        days = [i for i, (full, short) in enumerate(WEEKDAYS)
                if full in line or re.search(r"\b%s\b" % short, line)]
        if days and wd not in days:
            continue  # правило на другой день недели
        for m in re.finditer(r"(?:\bс|\bот)\s*%s\s*(?:ч\.?)?\s*(?:до|-|–)\s*%s" % (hhmm, hhmm), line):
            h1, m1, h2, m2 = int(m.group(1)), int(m.group(2) or 0), int(m.group(3)), int(m.group(4) or 0)
            if 5 <= h1 <= 23 and 6 <= h2 <= 23:
                earliest = max(earliest or 0, h1 * 60 + m1)
                latest = min(latest or 24 * 60, h2 * 60 + m2)
        for m in re.finditer(r"\bдо\s*%s(?!\s*(?:шт|кор|пач|уп|кг|руб|р\b|%%))" % hhmm, line):
            h, mm = int(m.group(1)), int(m.group(2) or 0)
            if 6 <= h <= 23 and (m.group(2) or h >= 9):  # «до 16», «до 16:00»; «до 5» — не время
                latest = min(latest or 24 * 60, h * 60 + mm)
        for m in re.finditer(r"\bпосле\s*%s" % hhmm, line):
            h, mm = int(m.group(1)), int(m.group(2) or 0)
            if 6 <= h <= 22:
                earliest = max(earliest or 0, h * 60 + mm)
    if earliest is None and latest is None:
        return None
    if earliest is not None and latest is not None and earliest >= latest:
        earliest = None  # противоречие («с 10:30» и «до 10:00») — важнее «до»
    return earliest, latest


def hm_to_min(text):
    """«11», «11:30», «11.30», «11-30» → минуты от полуночи; пусто/ошибка → None."""
    import re
    m = re.match(r"^\s*(\d{1,2})(?:\s*[:.\-]\s*(\d{2}))?\s*$", text or "")
    if not m or int(m.group(1)) > 23:
        return None
    return int(m.group(1)) * 60 + int(m.group(2) or 0)


def point_window(o, day, cache):
    """Часы приёма точки: сохранённые вручную — главнее, иначе из комментариев."""
    saved = cache.get("hours|" + geo_key(o))
    if saved is not None:
        w = (saved.get("from"), saved.get("to"))
        return None if w == (None, None) else w
    return time_window(" \n ".join([o.get("comment", ""), o.get("note", ""), o.get("address", ""),
                                     o.get("hours", "")]), day)


def fmt_hm(minutes):
    minutes = int(round(minutes))
    return "%d:%02d" % (minutes // 60 % 24, minutes % 60)


def eval_route(order, T, start, service, windows, back):
    """Проходит маршрут: время прибытия, опоздания. order — индексы точек 1..n (0 — цех)."""
    t, prev, late, eta = start, 0, 0.0, []
    for i in order:
        t += T[prev][i]
        w = windows[i]
        if w and w[0] and t < w[0]:
            t = w[0]  # приехали раньше открытия — ждём
        eta.append(t)
        if w and w[1] and t > w[1]:
            late += t - w[1]
        t += service
        prev = i
    if back:
        t += T[prev][0]
    return t - start + 1000 * late, late, eta


def solve_route(T, start, service, windows, back):
    """Лучший порядок объезда с учётом окон. До 11 точек — перебор (точно), больше — эвристика."""
    import itertools
    n = len(T) - 1
    pts = list(range(1, n + 1))
    if n <= 1:
        return pts
    if n <= 8:
        best = min(itertools.permutations(pts), key=lambda p: eval_route(p, T, start, service, windows, back)[0])
        return list(best)
    if n <= 11:  # динамика по подмножествам
        INF = float("inf")
        dp = {}
        for i in pts:
            dp[(1 << i, i)] = (eval_route([i], T, start, service, windows, False)[0], [i])
        for size in range(2, n + 1):
            new = {}
            for (mask, last), (cost, path) in dp.items():
                for j in pts:
                    if mask & (1 << j):
                        continue
                    p2 = path + [j]
                    c2 = eval_route(p2, T, start, service, windows, False)[0]
                    k = (mask | (1 << j), j)
                    if c2 < new.get(k, (INF,))[0]:
                        new[k] = (c2, p2)
            dp = new
        return min((v[1] for v in dp.values()), key=lambda p: eval_route(p, T, start, service, windows, back)[0])
    # эвристика: жадно (срочные — раньше), затем улучшения перестановками
    route, left, prev, t = [], set(pts), 0, start
    while left:
        def score(j):
            arr = t + T[prev][j]
            w = windows[j]
            urgency = (w[1] - arr) if (w and w[1]) else 10 ** 4
            return T[prev][j] + (0 if urgency > 60 else (60 - urgency) * 2)
        j = min(left, key=score)
        route.append(j)
        left.discard(j)
        t += T[prev][j] + service
        prev = j
    best = eval_route(route, T, start, service, windows, back)[0]
    improved = True
    while improved:
        improved = False
        for a in range(len(route) - 1):
            for b in range(a + 1, len(route)):
                for cand in (route[:a] + route[a:b + 1][::-1] + route[b + 1:],      # разворот отрезка
                             route[:a] + route[a + 1:b + 1] + [route[a]] + route[b + 1:]):  # перенос точки
                    c = eval_route(cand, T, start, service, windows, back)[0]
                    if c + 1e-6 < best:
                        route, best, improved = cand, c, True
    return route


def plan_route(cfg, day, orders, start_hm, service, back, progress=lambda t: None, manual=None):
    """Порядок объезда для выбранных отгрузок.
    → dict(order=[orders в порядке], eta={id: минуты}, windows={id: окно}, km, minutes, roads, missing=[orders])"""
    cache = geo_load()
    manual = manual or {}
    depot_addr = cfg.get("depot_address") or "проспект Космонавтов 34Г, Барнаул"
    depot = parse_coords(cfg.get("depot_coords", "")) or geocode(depot_addr, cache, key="цех", progress=progress)
    if not depot:
        raise Oops("Не нашёл на карте адрес цеха «%s». Впиши координаты в config.ini (depot_coords)." % depot_addr)
    cache["цех"] = {"ll": list(depot), "found": depot_addr}
    points, used, missing = [depot], [], []
    for o in orders:
        k = geo_key(o)
        if k in manual:  # поправили адрес в окне
            fix = manual[k]
            ll = parse_coords(fix) or geocode(fix, cache, progress=progress)
            if ll:
                cache[k] = {"ll": list(ll), "found": fix}
        ll = geocode(o["address"] or o["comment"], cache, key=k, progress=progress) \
            if (o["address"] or o["comment"] or k in cache) else None
        if ll:
            cache.setdefault(k, {"ll": list(ll), "found": o["address"]})
            points.append(ll)
            used.append(o)
        else:
            missing.append(o)
    geo_save(cache)
    if not used:
        return dict(order=list(orders), eta={}, windows={}, km=0, minutes=0, roads=False, missing=missing,
                    points={}, depot=depot)
    progress("Считаю дороги и время…")
    T, K, roads = road_matrix(points)
    windows = [None] + [point_window(o, day, cache) for o in used]
    start = int(start_hm[0]) * 60 + int(start_hm[1])
    progress("Подбираю порядок объезда…")
    seq = solve_route(T, start, service, windows, back)
    total, late, eta = eval_route(seq, T, start, service, windows, back)
    km = sum(K[a][b] for a, b in zip([0] + seq, seq + ([0] if back else [])))
    ordered = [used[i - 1] for i in seq]
    return dict(order=ordered + missing,
                eta={used[i - 1]["id"]: e for i, e in zip(seq, eta)},
                windows={used[i - 1]["id"]: windows[i] for i in seq},
                km=km, minutes=total - 1000 * late, late=late, roads=roads, missing=missing,
                points={used[i - 1]["id"]: points[i] for i in seq}, depot=depot)


def yandex_route_url(depot, pts):
    """Ссылка на маршрут в Яндекс Картах (открывается и в Навигаторе)."""
    return "https://yandex.ru/maps/?rtt=auto&rtext=" + "~".join("%.6f,%.6f" % p for p in [depot] + pts)


# ---------------------------------------------------------------- анализ клиентов
LEGAL = {"ип", "ооо", "оао", "зао", "пао", "ао", "нко", "тд", "ооо", "индивидуальный", "предприниматель",
         "общество", "ограниченной", "ответственностью"}


def _words(text):
    import re
    return re.findall(r"[a-zа-я0-9]+", yo(text))


def name_tokens(name):
    """Ключевые слова названия контрагента: без ИП/ООО, без скобок."""
    import re
    main = re.sub(r"\([^)]*\)", " ", name or "")
    toks = [w for w in _words(main) if w not in LEGAL and len(w) >= 3]
    return toks or [w for w in _words(name) if w not in LEGAL]


def line_matches(tokens, line):
    """Похожа ли строка списка на контрагента. Инициалы «К.А.» засчитываются за имя и отчество."""
    if not tokens:
        return False
    words = _words(line)
    ws = set(words)
    initials = [w for w in words if len(w) == 1]
    if tokens[0] not in ws:
        return False
    hit = 1
    for i, t in enumerate(tokens[1:], 1):
        if t in ws or (t[0] in initials):
            hit += 1
    return hit * 2 >= len(tokens)


def read_call_list(path):
    """Список обзвона из .xls/.xlsx: (название списка, [строки-наименования])."""
    rows = []
    if path.lower().endswith(".xls"):
        import xlrd
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    else:
        from openpyxl import load_workbook
        rows = [list(r) for r in load_workbook(path, read_only=True, data_only=True).active.iter_rows(values_only=True)]
    title, col, start = None, None, None
    for r, row in enumerate(rows[:10]):
        for c, v in enumerate(row):
            if isinstance(v, str) and yo(v) == "наименование":
                col, start = c, r + 1
            elif isinstance(v, str) and v.strip() and title is None and col is None and yo(v) not in ("код", "дата"):
                title = v.strip()
        if col is not None:
            break
    if col is None:
        col, start = 0, 0
    names = [str(row[col]).strip() for row in rows[start:] if col < len(row) and str(row[col] or "").strip()]
    return title or os.path.splitext(os.path.basename(path))[0], names


def fetch_period(api, d0, d1, progress=lambda t: None):
    rows, offset = [], 0
    while True:
        progress("Загружаю отгрузки: %d…" % len(rows))
        r = api._check(api.s.get(BASE + "/entity/demand", params={
            "filter": "moment>=%s 00:00:00;moment<=%s 23:59:59" % (d0.strftime("%Y-%m-%d"), d1.strftime("%Y-%m-%d")),
            "expand": "agent", "limit": 100, "offset": offset, "order": "moment,asc"}, timeout=120))
        batch = r.json().get("rows", [])
        rows.extend(batch)
        if len(batch) < 100:
            return rows
        offset += 100


def analyze_clients(api, cfg, months, list_paths, min_orders=3, progress=lambda t: None):
    """Отчёт: кто и как часто заказывает, кто не в списках обзвона, кому пора звонить."""
    from statistics import median
    today = datetime.now().date()
    d0 = today - timedelta(days=int(months * 30.4))
    demands = fetch_period(api, d0, today, progress)
    progress("Считаю…")

    lists = []
    for pth in list_paths:
        try:
            lists.append(read_call_list(pth))
        except Exception as e:
            raise Oops("Не получилось прочитать список «%s»: %s" % (os.path.basename(pth), e))

    clients = {}
    for d in demands:
        ag = d.get("agent") or {}
        cid = ag.get("id") or href_id(ag)
        c = clients.setdefault(cid, {"name": ag.get("name", "?"), "dates": set(), "sum": 0.0, "docs": 0,
                                     "route": attr_value(ag, "маршрут"), "phone": ag.get("phone", "")})
        c["dates"].add(datetime.strptime(d["moment"][:10], "%Y-%m-%d").date())
        c["sum"] += (d.get("sum") or 0) / 100.0
        c["docs"] += 1

    used_lines = set()
    table = []
    for c in clients.values():
        dates = sorted(c["dates"])
        n = len(dates)
        gaps = [(b - a).days for a, b in zip(dates, dates[1:])]
        every = round(median(gaps)) if gaps else None
        nxt = dates[-1] + timedelta(days=every) if every else None
        toks = name_tokens(c["name"])
        found = []
        for title, names in lists:
            for i, line in enumerate(names):
                if line_matches(toks, line):
                    found.append((title, line))
                    used_lines.add((title, i))
        table.append({
            "Клиент": c["name"], "Маршрут в МоёмСкладе": c["route"],
            "В списке": ", ".join(sorted({t for t, _ in found})) or "нет",
            "Совпало со строкой": "; ".join(l for _, l in found[:3]),
            "Заказов (дней)": n, "Первый": dates[0], "Последний": dates[-1],
            "Раз в, дней": every, "Следующий ожидается": nxt,
            "Просрочка, дней": (today - nxt).days if nxt and nxt < today else 0,
            "Сумма за период": round(c["sum"], 2), "Средний чек": round(c["sum"] / c["docs"], 2),
            "Телефон": c["phone"],
        })

    table.sort(key=lambda x: (-x["Заказов (дней)"], x["Клиент"]))
    not_listed = [t for t in table if t["В списке"] == "нет" and t["Заказов (дней)"] >= min_orders]
    call_now = sorted([t for t in table if t["Заказов (дней)"] >= min_orders and t["Просрочка, дней"] > 0],
                      key=lambda x: -x["Просрочка, дней"])
    silent = [{"Список": title, "Строка списка": line}
              for title, names in lists for i, line in enumerate(names) if (title, i) not in used_lines]

    wb = Workbook()
    def sheet(ws, rows, cols, widths):
        from openpyxl.styles import PatternFill
        head_fill = PatternFill("solid", fgColor="E8F0FE")
        for j, (col, w) in enumerate(zip(cols, widths), 1):
            c = ws.cell(row=1, column=j, value=col)
            c.font = Font(name="Arial", size=10, bold=True)
            c.fill = head_fill
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[c.column_letter].width = w
        for i, row in enumerate(rows, 2):
            for j, col in enumerate(cols, 1):
                v = row.get(col)
                c = ws.cell(row=i, column=j, value=v)
                c.font = Font(name="Arial", size=10)
                if hasattr(v, "year"):
                    c.number_format = "DD.MM.YYYY"
                elif isinstance(v, (int, float)) and ("Сумма" in col or "чек" in col):
                    c.number_format = "#,##0.00"
        ws.freeze_panes = "B2"
        if rows:
            ws.auto_filter.ref = ws.dimensions

    cols = ["Клиент", "Маршрут в МоёмСкладе", "Заказов (дней)", "Раз в, дней", "Последний",
            "Следующий ожидается", "Просрочка, дней", "Сумма за период", "Средний чек", "Телефон"]
    w = [42, 14, 10, 9, 12, 13, 10, 14, 12, 22]
    ws = wb.active
    ws.title = "Не в списках"
    sheet(ws, not_listed, cols, w)
    sheet(wb.create_sheet("Пора звонить"), call_now, ["Клиент", "В списке"] + cols[1:], [42, 12] + w[1:])
    sheet(wb.create_sheet("Все клиенты"), table,
          ["Клиент", "В списке", "Совпало со строкой"] + cols[1:4] + ["Первый"] + cols[4:],
          [42, 12, 40] + w[1:4] + [12] + w[4:])
    sheet(wb.create_sheet("Из списков без заказов"), silent, ["Список", "Строка списка"], [14, 60])

    os.makedirs(cfg["out_dir"], exist_ok=True)
    path = free_path(cfg, "Анализ клиентов %s (%d мес)" % (today.strftime("%d.%m.%Y"), months))
    wb.save(path)
    return path, {"клиентов": len(table), "не в списках": len(not_listed), "пора звонить": len(call_now),
                  "без заказов": len(silent), "отгрузок": len(demands)}


# ---------------------------------------------------------------- запуск
def parse_day(arg):
    if not arg:
        return datetime.now()
    a = arg.strip().lower()
    if a in ("вчера", "y"):
        return datetime.now() - timedelta(days=1)
    for f in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d.%m"):
        try:
            d = datetime.strptime(a, f)
            return d.replace(year=datetime.now().year) if f == "%d.%m" else d
        except ValueError:
            pass
    raise Oops("Не понял дату: %s (пример: 22.09.2026)" % arg)


def default_name(day, kind="Доставка", extra=""):
    """«Доставка 20.10.2026», «Командировка 20.10.2026 Рубцовск»."""
    return " ".join(x for x in [kind, day.strftime("%d.%m.%Y"), (extra or "").strip()] if x)


def looks_like_trip(orders, cfg):
    """Командировка по умолчанию — если среди отмеченных есть ТОРГСЕРВИС (клиенты из ttn_for)."""
    return any(o.get("ttn") for o in orders)


def safe_name(name):
    import re
    name = re.sub(r'[\\/:*?"<>|]+', " ", name or "").strip().rstrip(".")
    return " ".join(name.split()) or "Маршрутный лист"


def day_dir(cfg, day):
    """Папка для файлов одного дня отгрузок: <папка программы>/2026-09-29"""
    return os.path.join(cfg["out_dir"], day.strftime("%Y-%m-%d"))


def free_path(cfg, name, day=None, ext=".xlsx"):
    """Путь для нового файла (в папке дня, если указан день): если имя занято — «имя (2)», «имя (3)»…"""
    folder = day_dir(cfg, day) if day else cfg["out_dir"]
    base = safe_name(name)
    path = os.path.join(folder, base + ext)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, "%s (%d)%s" % (base, n, ext))
        n += 1
    return path


def tidy_old_files(cfg):
    """Раскладывает старые «Маршрутный лист ДД.ММ.ГГГГ…xlsx» и «Документы ДД.ММ.ГГГГ…pdf»
    из общей папки по папкам дней. Открытые в Excel файлы пропускает — перенесутся в следующий раз."""
    import re
    import filecmp
    root = cfg["out_dir"]
    moved = 0
    try:
        names = os.listdir(root)
    except OSError:
        return 0
    pat = re.compile(r"^(Маршрутный лист|Документы|Доставка|Командировка) (\d{2})\.(\d{2})\.(\d{4}).*\.(xlsx|pdf)$", re.I)
    for n in names:
        m = pat.match(n)
        src = os.path.join(root, n)
        if not m or n.startswith("~$") or not os.path.isfile(src):
            continue
        folder = os.path.join(root, "%s-%s-%s" % (m.group(4), m.group(3), m.group(2)))
        base, ext = os.path.splitext(n)
        try:
            os.makedirs(folder, exist_ok=True)
            same = [f for f in os.listdir(folder) if f.startswith(base) and f.endswith(ext)
                    and filecmp.cmp(src, os.path.join(folder, f), shallow=False)]
            if same:  # такой же файл уже в папке дня — старую копию просто убираем
                os.remove(src)
                continue
            dst = os.path.join(folder, n)
            k = 2
            while os.path.exists(dst):
                dst = os.path.join(folder, "%s (%d)%s" % (base, k, ext))
                k += 1
            os.rename(src, dst)  # только перенос, без копирования: открытый файл просто пропустим
            moved += 1
        except OSError:
            pass  # файл открыт в Excel — в следующий раз
    dedupe_day_folders(root)
    return moved


def dedupe_day_folders(root):
    """Убирает одинаковые копии «имя (2).xlsx», «имя (3).xlsx»…, если они точь-в-точь как «имя.xlsx»."""
    import re
    import filecmp
    for d in os.listdir(root):
        folder = os.path.join(root, d)
        if not (re.match(r"^\d{4}-\d{2}-\d{2}$", d) and os.path.isdir(folder)):
            continue
        for f in os.listdir(folder):
            m = re.match(r"^(.*) \(\d+\)(\.(xlsx|pdf))$", f)
            if not m:
                continue
            orig = os.path.join(folder, m.group(1) + m.group(2))
            dup = os.path.join(folder, f)
            try:
                if os.path.exists(orig) and filecmp.cmp(orig, dup, shallow=False):
                    os.remove(dup)
            except OSError:
                pass


def production_rows(orders):
    """Сводка для производства (как внизу маршрутного листа): [(товар, «2 уп (24 шт)», пометки)].
    Складываем в штуках (базовой единице) — неважно, чем оформили реализацию, — потом переводим в упаковки."""
    prod = OrderedDict()
    for o in orders:
        for g in o["goods"]:
            if not g["name"]:
                continue
            rec = prod.setdefault(g["name"], {"qty": 0.0, "per": None, "pu": "уп", "bu": "шт", "marks": []})
            if "qty" in g:
                rec["qty"] += g["qty"]
                if g.get("per") and not rec["per"]:
                    rec["per"], rec["pu"], rec["bu"] = g["per"], g["pu"], g["bu"]
                elif not rec["per"]:
                    rec["bu"] = g.get("bu", "шт")
            else:
                rec["qty"] += g["packs"] or 0
            for m in filter(None, (g.get("marks") or "").split(", ")):
                if m not in rec["marks"]:
                    rec["marks"].append(m)
    return [(name, qty_text(r["qty"], r["per"], r["pu"], r["bu"]), ", ".join(r["marks"])) for name, r in prod.items()]


def build_loading_xlsx(orders, path):
    """«Загрузка» для начальника производства: что водитель забрал со склада, под подпись."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Загрузка"
    ws.column_dimensions["A"].width = 45.7
    ws.column_dimensions["B"].width = 19
    ws.column_dimensions["C"].width = 24.4
    r = 1
    for name, qty, marks in production_rows(orders):
        cell(ws, "A%d" % r, name)
        qty_cell(ws, "B%d" % r, qty)
        if marks:
            cell(ws, "C%d" % r, marks, BOLD, border=False, align=CENTER)
        lines = max(1, -(-len(name) // 34), -(-len(str(qty)) // 24))  # сколько строк займёт текст
        ws.row_dimensions[r].height = 16.5 * lines + (3 if lines > 1 else 0)
        r += 1
    r += 1
    cell(ws, "A%d" % r, "Водитель", border=False)
    ws.row_dimensions[r].height = 23.25
    cell(ws, "A%d" % (r + 1), "Дата", border=False)
    ws.row_dimensions[r + 1].height = 28.5
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    wb.save(path)


def order_path(cfg, day):
    return os.path.join(day_dir(cfg, day), "порядок.json")


def remember_order(cfg, day, ids):
    """Запоминает порядок отгрузок из маршрутного листа — по нему потом идёт печать."""
    import json
    old = []
    try:
        with open(order_path(cfg, day), encoding="utf-8") as f:
            old = json.load(f)
    except Exception:
        pass
    new = list(ids)
    if set(new) & set(old):  # тот же лист переделали — новый порядок главнее
        res = new + [x for x in old if x not in new]
    else:  # вторая машина / второй лист за день — после первого
        res = old + new
    try:
        os.makedirs(day_dir(cfg, day), exist_ok=True)
        with open(order_path(cfg, day), "w", encoding="utf-8") as f:
            json.dump(res, f)
    except OSError:
        pass


def sort_like_sheet(cfg, day, orders):
    """Сортирует отгрузки в порядке маршрутных листов этого дня (остальные — в конце)."""
    import json
    try:
        with open(order_path(cfg, day), encoding="utf-8") as f:
            pos = {x: i for i, x in enumerate(json.load(f))}
    except Exception:
        return orders
    idx = {id(o): i for i, o in enumerate(orders)}
    orders.sort(key=lambda o: (0, pos[o["id"]]) if o["id"] in pos else (1, idx[id(o)]))
    return orders


def save_sheet(orders, cfg, day, name=None, route_url=None):
    """Сохраняет маршрутный лист по выбранным отгрузкам. Возвращает (путь, доставок, наличка, подсказка)."""
    if not orders:
        raise Oops("Не выбрано ни одной отгрузки.")
    name = safe_name(name or default_name(day))
    path = free_path(cfg, name, day)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        build_xlsx(orders, day, cfg, path, title=name, route_url=route_url)
        remember_order(cfg, day, [o["id"] for o in orders])
        load_path = path[:-5] + " — загрузка.xlsx"
        build_loading_xlsx(orders, load_path)
    except PermissionError:
        raise Oops("Не получилось сохранить «%s» — файл открыт в Excel?" % os.path.basename(path))
    cash = [o for o in orders if o["cash"]]
    hint = ""
    if not cash:  # подсказка, если наличка не нашлась
        hint = "Налички нет. Организации в отгрузках: %s. Если «ГК» записана иначе — поправь cash_org в config.ini." % (
            ", ".join(sorted({o["org"] for o in orders if o["org"]})) or "не указаны")
    return path, len(orders), sum(o["sum"] for o in cash), hint, sum(o["sum"] for o in orders), load_path


def generate(api, cfg, day):
    """Маршрутный лист за день по всем отгрузкам (для запуска из консоли)."""
    return save_sheet(load_orders(api, cfg, day), cfg, day)


def open_file(path):
    if os.name == "nt":
        os.startfile(path)
    else:
        import subprocess
        try:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        except Exception:
            pass


def make_api(cfg):
    return Api(cfg)


# ---------------------------------------------------------------- окно
def money(x):
    return "{:,.2f}".format(x).replace(",", " ")


# Цвета интерфейса
C = {
    "bg": "#F3F5F8", "card": "#FFFFFF", "line": "#E3E7EE", "text": "#1F2933", "muted": "#6B7785",
    "accent": "#1F6FEB", "accent_hover": "#1858BD", "accent_soft": "#E8F0FE",
    "ok": "#1A7F37", "ok_soft": "#E6F4EA", "warn": "#9A6700", "warn_soft": "#FFF4D6",
    "err": "#C62828", "err_soft": "#FDECEC", "row_alt": "#F8FAFC",
}
UI = "Segoe UI" if os.name == "nt" else ("Helvetica Neue" if sys.platform == "darwin" else "DejaVu Sans")


MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август",
          "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]


def calendar_popup(root, anchor, date_var, F):
    """Календарь под кнопкой: клик по дню ставит дату."""
    import calendar
    import customtkinter as ctk

    try:
        sel = datetime.strptime(date_var.get().strip(), "%d.%m.%Y").date()
    except ValueError:
        sel = datetime.now().date()
    today = datetime.now().date()
    state = {"y": sel.year, "m": sel.month}

    pop = ctk.CTkToplevel(root, fg_color=C["card"])
    pop.overrideredirect(True)
    pop.transient(root)
    x, y = anchor.winfo_rootx() - 140, anchor.winfo_rooty() + anchor.winfo_height() + 6
    pop.geometry("+%d+%d" % (x, y))
    frame = ctk.CTkFrame(pop, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["line"])
    frame.pack(padx=1, pady=1)
    top = ctk.CTkFrame(frame, fg_color="transparent")
    top.pack(fill="x", padx=10, pady=(10, 4))
    title = ctk.CTkLabel(top, text="", font=F(14, True), text_color=C["text"])
    grid = ctk.CTkFrame(frame, fg_color="transparent")
    grid.pack(padx=10, pady=(0, 10))
    nav = dict(width=32, height=32, corner_radius=8, font=F(14, True), fg_color="transparent",
               hover_color=C["line"], text_color=C["text"])

    def move(d):
        m = state["m"] + d
        state["y"] += (m - 1) // 12
        state["m"] = (m - 1) % 12 + 1
        draw()

    ctk.CTkButton(top, text="‹", command=lambda: move(-1), **nav).pack(side="left")
    ctk.CTkButton(top, text="›", command=lambda: move(1), **nav).pack(side="right")
    title.pack(side="left", expand=True)

    bind_ids = {}

    def close(*_):
        for ev, bid in bind_ids.items():
            try:
                root.unbind(ev, bid)
            except Exception:
                pass
        bind_ids.clear()
        if pop.winfo_exists():
            pop.destroy()

    def choose(d):
        date_var.set(d.strftime("%d.%m.%Y"))
        close()

    def draw():
        for w in grid.winfo_children():
            w.destroy()
        title.configure(text="%s %d" % (MONTHS[state["m"] - 1], state["y"]))
        for c, wd in enumerate(["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]):
            ctk.CTkLabel(grid, text=wd, width=36, font=F(11, True),
                         text_color=C["err"] if c >= 5 else C["muted"]).grid(row=0, column=c, pady=(0, 2))
        for r, week in enumerate(calendar.Calendar().monthdatescalendar(state["y"], state["m"]), 1):
            for c, d in enumerate(week):
                other = d.month != state["m"]
                is_sel, is_today = d == sel, d == today
                ctk.CTkButton(
                    grid, text=str(d.day), width=36, height=32, corner_radius=8, font=F(12, is_sel or is_today),
                    fg_color=C["accent"] if is_sel else "transparent",
                    hover_color=C["accent_hover"] if is_sel else C["accent_soft"],
                    text_color="#FFFFFF" if is_sel else ("#B8C1CC" if other else (C["accent"] if is_today else C["text"])),
                    border_width=1 if is_today and not is_sel else 0, border_color=C["accent"],
                    command=lambda d=d: choose(d)).grid(row=r, column=c, padx=1, pady=1)
        ctk.CTkButton(grid, text="Сегодня", height=28, corner_radius=8, font=F(12), fg_color=C["accent_soft"],
                      hover_color=C["line"], text_color=C["accent"],
                      command=lambda: choose(today)).grid(row=8, column=0, columnspan=7, sticky="ew", pady=(6, 0))

    draw()
    # закрыть по клику в главном окне или Esc
    pop.bind("<Escape>", close)
    root.after(50, lambda: bind_ids.update({
        "<Button-1>": root.bind("<Button-1>", close, add="+"),
        "<Escape>": root.bind("<Escape>", close, add="+")}))



def fix_clipboard_keys(root):
    """Ctrl+V/C/X/A работают при русской раскладке + меню по правому клику во всех полях ввода."""
    import tkinter as tk
    codes = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>"}

    def on_key(e):
        if e.keysym.lower() in ("v", "c", "x", "a"):
            return None  # латинская раскладка — Tk справится сам
        if e.keycode in codes:
            e.widget.event_generate(codes[e.keycode])
            return "break"
        if e.keycode == 65:  # Ctrl+A
            try:
                e.widget.select_range(0, "end")
                e.widget.icursor("end")
            except Exception:
                pass
            return "break"
        return None

    def on_right(e):
        w = e.widget
        if not isinstance(w, (tk.Entry, tk.Text)):
            return
        m = tk.Menu(root, tearoff=0)
        m.add_command(label="Вставить", command=lambda: w.event_generate("<<Paste>>"))
        m.add_command(label="Копировать", command=lambda: w.event_generate("<<Copy>>"))
        m.add_command(label="Вырезать", command=lambda: w.event_generate("<<Cut>>"))
        m.tk_popup(e.x_root, e.y_root)

    root.bind_all("<Control-KeyPress>", on_key, add="+")
    root.bind_all("<Button-3>", on_right, add="+")


def settings_dialog(root, F, first_run=False, on_done=None):
    """Окно настройки: вставить токен МоегоСклада, проверить, сохранить."""
    import threading
    import customtkinter as ctk

    win = ctk.CTkToplevel(root, fg_color=C["bg"])
    win.title("Настройка" if not first_run else "Первый запуск")
    win.transient(root)
    win.resizable(False, False)
    box = ctk.CTkFrame(win, fg_color="transparent")
    box.pack(padx=24, pady=20)
    ctk.CTkLabel(box, text="Подключение к МоемуСкладу" if not first_run else "Добро пожаловать!",
                 font=F(20, True), text_color=C["text"]).pack(anchor="w")
    ctk.CTkLabel(box, justify="left", font=F(12), text_color=C["muted"], text=(
        "Нужен твой личный токен МоегоСклада.\n"
        "Где взять: МойСклад → Настройки → Токены → «Создать токен».\n"
        "Если такого пункта нет — попроси администратора дать тебе доступ к API.\n"
        "Скопируй токен целиком и вставь сюда.")).pack(anchor="w", pady=(4, 12))
    tok = ctk.StringVar(value="" if not has_token() else load_config()["token"])
    row = ctk.CTkFrame(box, fg_color="transparent")
    row.pack(anchor="w", fill="x")
    entry = ctk.CTkEntry(row, textvariable=tok, width=350, height=40, font=F(13), corner_radius=10, show="•",
                         border_color=C["line"], fg_color=C["card"])
    entry.pack(side="left")

    def paste():
        try:
            tok.set(win.clipboard_get().strip())
        except Exception:
            msg.configure(text="В буфере обмена пусто — сначала скопируй токен в МоёмСкладе.", text_color=C["err"])
    ctk.CTkButton(row, text="Вставить", width=100, height=40, corner_radius=10, font=F(13, True),
                  fg_color=C["accent_soft"], hover_color=C["line"], text_color=C["accent"],
                  command=paste).pack(side="left", padx=(10, 0))
    shortcut = ctk.BooleanVar(value=first_run and os.name == "nt")
    if os.name == "nt":
        ctk.CTkCheckBox(box, text="Создать ярлык на рабочем столе", variable=shortcut, font=F(13),
                        fg_color=C["accent"], hover_color=C["accent_hover"]).pack(anchor="w", pady=(12, 0))
    msg = ctk.CTkLabel(box, text="", font=F(13), justify="left", wraplength=460)
    msg.pack(anchor="w", pady=(10, 0))
    bar = ctk.CTkFrame(box, fg_color="transparent")
    bar.pack(fill="x", pady=(12, 0))

    def save():
        t = tok.get().strip()
        if not t:
            msg.configure(text="Вставь токен.", text_color=C["err"])
            return
        ok_btn.configure(state="disabled")
        msg.configure(text="Проверяю в МоёмСкладе…", text_color=C["muted"])

        def work():
            try:
                who = check_token(t)
                save_token(t)
                made = shortcut.get() and make_shortcut()
                def done():
                    msg.configure(text="✓  Готово! Токен сотрудника: %s%s" % (
                        who, "\nЯрлык «Маршрутный лист» на рабочем столе." if made else ""), text_color=C["ok"])
                    ok_btn.configure(text="Закрыть", state="normal", command=close)
                win.after(0, done)
            except Oops as e:
                err = str(e)
                win.after(0, lambda: (msg.configure(text=err, text_color=C["err"]),
                                      ok_btn.configure(state="normal")))
        threading.Thread(target=work, daemon=True).start()

    def close():
        win.destroy()
        if on_done:
            on_done()

    ok_btn = ctk.CTkButton(bar, text="Проверить и сохранить", width=200, height=42, corner_radius=12,
                           font=F(14, True), fg_color=C["accent"], hover_color=C["accent_hover"], command=save)
    ok_btn.pack(side="right")
    ctk.CTkButton(bar, text="Закрыть" if first_run else "Отмена", width=90, height=42, corner_radius=12,
                  font=F(13), fg_color=C["card"], hover_color=C["line"], text_color=C["text"], border_width=1,
                  border_color=C["line"], command=win.destroy).pack(side="right", padx=8)
    # окно — поверх главного и с курсором в поле
    win.after(150, lambda: (win.lift(), win.focus_force(), entry.focus_set()))
    return win


def analysis_dialog(root, F, run):
    """Окно «Анализ клиентов»: период, минимум заказов, списки обзвона."""
    import customtkinter as ctk
    from tkinter import filedialog

    win = ctk.CTkToplevel(root, fg_color=C["bg"])
    win.title("Анализ клиентов")
    win.transient(root)
    win.resizable(False, False)
    box = ctk.CTkFrame(win, fg_color="transparent")
    box.pack(padx=24, pady=20)
    ctk.CTkLabel(box, text="Анализ клиентов", font=F(20, True), text_color=C["text"]).pack(anchor="w")
    ctk.CTkLabel(box, justify="left", font=F(12), text_color=C["muted"], text=(
        "Кто и как часто заказывает, кого нет в списках обзвона и кому пора звонить.\n"
        "Данные берутся из отгрузок МоегоСклада за выбранный период.")).pack(anchor="w", pady=(4, 14))

    card = ctk.CTkFrame(box, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["line"])
    card.pack(fill="x")
    ctk.CTkLabel(card, text="Период", font=F(12, True), text_color=C["muted"]).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 6))
    period = ctk.CTkSegmentedButton(card, values=["3 мес", "6 мес", "12 мес"], height=36, font=F(13),
                                    selected_color=C["accent"], selected_hover_color=C["accent_hover"],
                                    unselected_color=C["bg"], unselected_hover_color=C["line"], text_color=C["text"])
    period.set(get_setting("analysis", "period", "6 мес"))
    period.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=(14, 6))
    ctk.CTkLabel(card, text="Постоянный —\nот заказов", font=F(12, True), text_color=C["muted"], justify="left").grid(
        row=1, column=0, sticky="w", padx=16, pady=(6, 14))
    min_var = ctk.StringVar(value=get_setting("analysis", "min_orders", "3"))
    ctk.CTkEntry(card, textvariable=min_var, width=60, height=34, justify="center", font=F(13),
                 border_color=C["line"], fg_color=C["bg"]).grid(row=1, column=1, sticky="w", pady=(6, 14))

    ctk.CTkLabel(box, text="Списки обзвона (Excel)", font=F(12, True), text_color=C["muted"]).pack(anchor="w", pady=(14, 4))
    paths = [x for x in get_setting("analysis", "lists", "").split("|") if x.strip()]
    lst = ctk.CTkLabel(box, text="", font=F(13), justify="left", anchor="w", text_color=C["text"], wraplength=440)
    lst.pack(anchor="w", fill="x")

    def show():
        lst.configure(text="\n".join("•  " + os.path.basename(x) + ("" if os.path.exists(x) else "  (файл не найден)")
                                      for x in paths) or "не выбраны — отчёт посчитает частоту заказов без сверки")
    show()
    row = ctk.CTkFrame(box, fg_color="transparent")
    row.pack(anchor="w", pady=(6, 0))
    small = dict(height=32, corner_radius=8, font=F(12), fg_color=C["card"], hover_color=C["line"],
                 text_color=C["text"], border_width=1, border_color=C["line"])

    def add():
        got = filedialog.askopenfilenames(parent=win, title="Списки обзвона",
                                          filetypes=[("Excel", "*.xls *.xlsx"), ("Все файлы", "*.*")])
        for g in got:
            if g not in paths:
                paths.append(g)
        show()

    def clear():
        paths.clear()
        show()
    ctk.CTkButton(row, text="Добавить файлы…", width=140, command=add, **small).pack(side="left")
    ctk.CTkButton(row, text="Очистить", width=90, command=clear, **small).pack(side="left", padx=8)

    bar = ctk.CTkFrame(box, fg_color="transparent")
    bar.pack(fill="x", pady=(18, 0))

    def go():
        try:
            m = max(1, int(min_var.get()))
        except ValueError:
            m = 3
        save_setting("analysis", "lists", "|".join(paths))
        save_setting("analysis", "period", period.get())
        save_setting("analysis", "min_orders", str(m))
        win.destroy()
        run(int(period.get().split()[0]), [x for x in paths if os.path.exists(x)], m)

    ctk.CTkButton(bar, text="Сформировать отчёт", width=190, height=42, corner_radius=12, font=F(14, True),
                  fg_color=C["accent"], hover_color=C["accent_hover"], command=go).pack(side="right")
    ctk.CTkButton(bar, text="Отмена", width=90, height=42, corner_radius=12, font=F(13), fg_color=C["card"],
                  hover_color=C["line"], text_color=C["text"], border_width=1, border_color=C["line"],
                  command=win.destroy).pack(side="right", padx=8)
    win.after(150, lambda: (win.lift(), win.focus_force()))
    return win


def gui():
    import threading
    import customtkinter as ctk

    ctk.set_appearance_mode("light")
    root = ctk.CTk(fg_color=C["bg"])
    fix_clipboard_keys(root)
    root.title("Маршрутный лист — Папирус Алтай")
    root.resizable(False, False)
    F = lambda size, bold=False: ctk.CTkFont(family=UI, size=size, weight="bold" if bold else "normal")

    wrap = ctk.CTkFrame(root, fg_color="transparent")
    wrap.pack(padx=24, pady=22)

    # --- шапка
    ctk.CTkLabel(wrap, text="ПАПИРУС АЛТАЙ", font=F(11, True), text_color=C["accent"]).pack(anchor="w")
    ctk.CTkButton(wrap, text="Настройки", width=80, height=26, corner_radius=8, font=F(12),
                  fg_color="transparent", hover_color=C["line"], text_color=C["muted"],
                  command=lambda: settings_dialog(root, F)).place(relx=1.0, x=0, y=-4, anchor="ne")
    ctk.CTkLabel(wrap, text="Маршрутный лист и документы", font=F(22, True),
                 text_color=C["text"]).pack(anchor="w", pady=(0, 2))
    ctk.CTkLabel(wrap, text="Отгрузки берутся из МоегоСклада за выбранную дату  ·  версия %s" % VERSION,
                 font=F(12), text_color=C["muted"]).pack(anchor="w", pady=(0, 16))

    # --- карточка с датой
    card = ctk.CTkFrame(wrap, fg_color=C["card"], corner_radius=14, border_width=1, border_color=C["line"])
    card.pack(fill="x")
    ctk.CTkLabel(card, text="Дата отгрузок", font=F(12, True), text_color=C["muted"]).pack(
        anchor="w", padx=18, pady=(14, 6))
    row = ctk.CTkFrame(card, fg_color="transparent")
    row.pack(fill="x", padx=18, pady=(0, 16))
    date_var = ctk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
    ctk.CTkEntry(row, textvariable=date_var, width=130, height=40, font=F(16, True), justify="center",
                 corner_radius=10, border_color=C["line"], fg_color=C["bg"]).pack(side="left")
    cal_btn = ctk.CTkButton(row, text="▾", width=40, height=40, corner_radius=10, font=F(16, True),
                            fg_color=C["bg"], hover_color=C["line"], text_color=C["text"],
                            border_width=1, border_color=C["line"])
    cal_btn.pack(side="left", padx=(6, 0))
    cal_btn.configure(command=lambda: calendar_popup(root, cal_btn, date_var, F))

    def pick(label):
        d = {"Вчера": -1, "Сегодня": 0, "Завтра": 1}[label]
        date_var.set((datetime.now() + timedelta(days=d)).strftime("%d.%m.%Y"))

    seg = ctk.CTkSegmentedButton(row, values=["Вчера", "Сегодня", "Завтра"], command=pick, height=40,
                                 font=F(13), corner_radius=10, selected_color=C["accent"],
                                 selected_hover_color=C["accent_hover"], unselected_color=C["bg"],
                                 unselected_hover_color=C["line"], text_color=C["text"])
    seg.set("Сегодня")
    seg.pack(side="left", padx=(12, 0))
    date_var.trace_add("write", lambda *_: seg.set(
        {(datetime.now() + timedelta(days=d)).strftime("%d.%m.%Y"): n
         for n, d in (("Вчера", -1), ("Сегодня", 0), ("Завтра", 1))}.get(date_var.get(), "")))

    # --- кнопки
    btn = ctk.CTkButton(wrap, text="Сформировать маршрутный лист…", height=52, corner_radius=12,
                        font=F(15, True), fg_color=C["accent"], hover_color=C["accent_hover"])
    btn.pack(fill="x", pady=(16, 10))
    btn2 = ctk.CTkButton(wrap, text="Печать УПД и счетов…", height=52, corner_radius=12, font=F(15, True),
                         fg_color=C["card"], hover_color=C["accent_soft"], text_color=C["accent"],
                         border_width=2, border_color=C["accent"])
    btn2.pack(fill="x")
    btn3 = ctk.CTkButton(wrap, text="Анализ клиентов…", height=40, corner_radius=12, font=F(13, True),
                         fg_color="transparent", hover_color=C["accent_soft"], text_color=C["accent"])
    btn3.pack(fill="x", pady=(8, 0))

    # --- статус
    stat = ctk.CTkFrame(wrap, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["line"])
    stat_text = ctk.CTkLabel(stat, text="", font=F(13), justify="left", anchor="w", wraplength=400)
    stat_text.pack(fill="x", padx=16, pady=(12, 4))
    bar = ctk.CTkProgressBar(stat, mode="indeterminate", height=6, progress_color=C["accent"], fg_color=C["line"])
    reopen = ctk.CTkButton(stat, text="Открыть файл ещё раз", height=30, width=10, font=F(12),
                           fg_color="transparent", hover_color=C["accent_soft"], text_color=C["accent"])
    stop_btn = ctk.CTkButton(stat, text="Отменить", height=32, width=110, corner_radius=8, font=F(13, True),
                             fg_color=C["card"], hover_color=C["err_soft"], text_color=C["err"],
                             border_width=1, border_color=C["err"])
    last = {"path": None}
    reopen.configure(command=lambda: last["path"] and open_file(last["path"]))

    def show_status(text, kind="info"):
        colors = {"info": (C["card"], C["text"]), "ok": (C["ok_soft"], C["ok"]), "err": (C["err_soft"], C["err"])}
        bg, fg = colors[kind]
        stat.configure(fg_color=bg, border_color=bg if kind != "info" else C["line"])
        stat_text.configure(text=text, text_color=fg)
        if not stat.winfo_ismapped():
            stat.pack(fill="x", pady=(16, 0))
        if kind == "info":
            reopen.pack_forget()
            bar.pack(fill="x", padx=16, pady=(4, 14))
            bar.start()
        else:
            bar.stop()
            bar.pack_forget()
            stop_btn.pack_forget()
            if kind == "ok" and last["path"]:
                reopen.pack(anchor="w", padx=8, pady=(0, 10))
            else:
                reopen.pack_forget()
                stat_text.pack_configure(pady=(12, 12))

    def say(text):
        root.after(0, lambda: stat_text.configure(text=text))

    def busy(on):
        for b in (btn, btn2, btn3):
            b.configure(state="disabled" if on else "normal")

    def finish(ok, msg, path=None, opener=None):
        busy(False)
        if ok and path:
            last["path"] = path
        show_status(msg, "ok" if ok else "err")
        if ok and path:
            (opener or open_file)(path)

    def background(job, then, cancel=None):
        """cancel — threading.Event: если передан, в плашке появляется кнопка «Отменить»."""
        def work():
            try:
                res = job()
                root.after(0, then, res)
            except Oops as e:
                root.after(0, finish, False, str(e))
            except Exception as e:
                root.after(0, finish, False, "Ошибка: %s" % e)
        busy(True)
        show_status("Загружаю отгрузки…")
        if cancel is not None:
            def do_cancel():
                cancel.set()
                stop_btn.configure(state="disabled", text="Отменяю…")
                stat_text.configure(text="Отменяю…")
            stop_btn.configure(state="normal", text="Отменить", command=do_cancel)
            stop_btn.pack(anchor="w", padx=16, pady=(0, 14), after=bar)
        threading.Thread(target=work, daemon=True).start()

    def ctx():
        cfg = load_config()
        return cfg, make_api(cfg), parse_day(date_var.get())

    # --- фоновая подготовка документов после маршрутного листа
    prep = {}
    prep_var = ctk.StringVar(value="")
    prep_lbl = ctk.CTkLabel(wrap, textvariable=prep_var, font=F(12), text_color=C["muted"], anchor="w")

    def start_prep(cfg, api, day, orders, ids):
        import threading
        same = prep.get("day") == day.date() and prep.get("pf")
        if not same:
            if prep.get("pf"):
                prep["pf"].close()
            prep.clear()
            prep.update(day=day.date(), pf=Prefetcher(api), ids=set(), api=api)
        prep["ids"] |= ids  # второй лист за тот же день — добавляем
        prep["ready"] = threading.Event()
        prep["orders"] = None
        prep["error"] = None

        def work(ev=prep["ready"], pf=prep["pf"], only=set(prep["ids"])):
            try:
                annotate(api, orders, only=only)
                pf.warm(orders, cfg)
                prep["orders"] = orders
            except Exception as e:
                prep["error"] = str(e)
            finally:
                ev.set()
        threading.Thread(target=work, daemon=True).start()
        if not prep_lbl.winfo_ismapped():
            prep_lbl.pack(anchor="w", pady=(8, 0), after=btn3)
        prep_tick()

    def prep_tick():
        pf = prep.get("pf")
        if not pf:
            prep_var.set("")
            prep_lbl.pack_forget()
            return
        if prep.get("error"):
            prep_var.set("Документы заранее подготовить не вышло — скачаются при печати.")
            return
        d, n = pf.progress()
        if prep.get("orders") is None:
            prep_var.set("Готовлю документы для печати…")
        elif d < n:
            prep_var.set("Документы для печати (%s): скачано %d из %d" % (prep["day"].strftime("%d.%m"), d, n))
        else:
            prep_var.set("Документы для печати (%s) готовы ✓" % prep["day"].strftime("%d.%m"))
        root.after(700, prep_tick)

    def drop_prep():
        if prep.get("pf"):
            prep["pf"].close()
        prep.clear()

    def sheet():
        def job():
            cfg, api, day = ctx()
            return cfg, api, day, load_orders(api, cfg, day)
        background(job, lambda res: choose_sheet(*res))

    def sheet_done(res):
        finish(True, "✓  Готово: %d доставок на %s ₽\nНаличка: %s ₽\n%s\n+ %s%s" % (
            res[1], money(res[4]), money(res[2]), os.path.basename(res[0]), os.path.basename(res[5]),
            "\n\n" + res[3] if res[3] else ""), res[0], lambda p: (open_file(p), open_file(res[5])))

    # --- окно выбора отгрузок для маршрутного листа
    def choose_sheet(cfg, api, day, orders):
        busy(False)
        bar.stop()
        stat.pack_forget()
        win = ctk.CTkToplevel(root, fg_color=C["bg"])
        win.title("Маршрутный лист — %s" % day.strftime("%d.%m.%Y"))
        win.transient(root)
        win.after(100, win.grab_set)
        box = ctk.CTkFrame(win, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=22, pady=18)
        ctk.CTkLabel(box, text="Что попадёт в маршрутный лист · %s" % day.strftime("%d.%m.%Y"),
                     font=F(20, True), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(box, text="Сними галочки с самовывоза и всего, что не едет с водителем.",
                     font=F(12), text_color=C["muted"]).pack(anchor="w", pady=(0, 12))

        widths = (40, 34, 64, 230, 330, 96, 100)
        head = ctk.CTkFrame(box, fg_color="transparent")
        head.pack(fill="x", padx=(10, 26))
        for i, (h, w) in enumerate(zip(["", "№", "Номер", "Контрагент", "Адрес", "Приезд", "Сумма"], widths)):
            ctk.CTkLabel(head, text=h, width=w, anchor="e" if i == 6 else "w", font=F(12, True),
                         text_color=C["muted"]).grid(row=0, column=i, padx=4)
        lst = ctk.CTkScrollableFrame(box, width=sum(widths) + 215, height=min(50 * len(orders), 420),
                                     fg_color=C["card"], corner_radius=14, border_width=1, border_color=C["line"])
        lst.pack(fill="both", expand=True, pady=(4, 0))

        counter = ctk.StringVar()
        keep = {}
        for o in orders:
            v = ctk.BooleanVar(value=True)
            keep[o["id"]] = v
        route = {"eta": {}, "windows": {}, "points": {}, "depot": None}
        gcache = geo_load()

        def edit_hours(o):
            """Окошко «Работают с … до …» для точки — запоминается для клиента."""
            dlg = ctk.CTkToplevel(win, fg_color=C["bg"])
            dlg.title("Часы приёма")
            dlg.transient(win)
            dlg.resizable(False, False)
            fb = ctk.CTkFrame(dlg, fg_color="transparent")
            fb.pack(padx=20, pady=16)
            ctk.CTkLabel(fb, text=o["agent"], font=F(15, True), text_color=C["text"], wraplength=360,
                         justify="left").pack(anchor="w")
            ctk.CTkLabel(fb, text="Когда принимают товар. Пусто — без ограничения. Запомнится для этого клиента.",
                         font=F(12), text_color=C["muted"], wraplength=360, justify="left").pack(anchor="w", pady=(2, 10))
            w = point_window(o, day, gcache) or (None, None)
            rowf = ctk.CTkFrame(fb, fg_color="transparent")
            rowf.pack(anchor="w")
            v_from, v_to = ctk.StringVar(value=fmt_hm(w[0]) if w[0] else ""), ctk.StringVar(value=fmt_hm(w[1]) if w[1] else "")
            for lbl, var in (("с", v_from), ("до", v_to)):
                ctk.CTkLabel(rowf, text=lbl, font=F(13), text_color=C["muted"]).pack(side="left", padx=(0, 6))
                ctk.CTkEntry(rowf, textvariable=var, width=70, height=34, justify="center", font=F(14),
                             border_color=C["line"], fg_color=C["card"]).pack(side="left", padx=(0, 14))
            note = ctk.CTkLabel(fb, text="", font=F(12), text_color=C["err"])
            note.pack(anchor="w", pady=(6, 0))

            def save():
                a, b = hm_to_min(v_from.get()), hm_to_min(v_to.get())
                if (v_from.get().strip() and a is None) or (v_to.get().strip() and b is None):
                    note.configure(text="Время пиши как 10:00 или 10")
                    return
                gcache["hours|" + geo_key(o)] = {"from": a, "to": b}
                geo_save(gcache)
                route["windows"][o["id"]] = None if (a, b) == (None, None) else (a, b)
                dlg.destroy()
                render()

            def reset():
                gcache.pop("hours|" + geo_key(o), None)
                geo_save(gcache)
                route["windows"].pop(o["id"], None)
                dlg.destroy()
                render()
            bb = ctk.CTkFrame(fb, fg_color="transparent")
            bb.pack(fill="x", pady=(12, 0))
            ctk.CTkButton(bb, text="Сохранить", width=120, height=36, corner_radius=10, font=F(13, True),
                          fg_color=C["accent"], hover_color=C["accent_hover"], command=save).pack(side="right")
            ctk.CTkButton(bb, text="Как в комментарии", width=150, height=36, corner_radius=10, font=F(12),
                          fg_color=C["card"], hover_color=C["line"], text_color=C["text"], border_width=1,
                          border_color=C["line"], command=reset).pack(side="right", padx=8)
            dlg.after(150, lambda: (dlg.lift(), dlg.focus_force()))

        def current_url():
            """Ссылка по текущему порядку (после стрелок тоже) и только отмеченным точкам."""
            pts = [route["points"][o["id"]] for o in orders if keep[o["id"]].get() and o["id"] in route["points"]]
            if not (pts and route["depot"]):
                return None
            if back_var.get():
                pts = pts + [route["depot"]]
            return yandex_route_url(route["depot"], pts)

        def recount(*_):
            chosen = [o for o in orders if keep[o["id"]].get()]
            counter.set("В листе: %d из %d · на %s ₽ · наличка %s ₽" % (
                len(chosen), len(orders), money(sum(o["sum"] for o in chosen)),
                money(sum(o["sum"] for o in chosen if o["cash"]))))
        for v in keep.values():
            v.trace_add("write", recount)

        def move(o, d):
            i = orders.index(o)
            j = i + d
            if 0 <= j < len(orders):
                orders[i], orders[j] = orders[j], orders[i]
                render()

        def render():
            for w in lst.winfo_children():
                w.destroy()
            arrow = dict(width=26, height=22, corner_radius=6, font=F(11), fg_color="transparent",
                         hover_color=C["line"], text_color=C["muted"])
            for r, o in enumerate(orders, 1):
                line = ctk.CTkFrame(lst, fg_color=C["row_alt"] if r % 2 == 0 else C["card"], corner_radius=8)
                line.pack(fill="x", pady=1)
                ctk.CTkCheckBox(line, text="", variable=keep[o["id"]], width=widths[0], corner_radius=6,
                                checkbox_width=20, checkbox_height=20, fg_color=C["accent"],
                                hover_color=C["accent_hover"], border_color="#B8C1CC").grid(row=0, column=0, padx=(8, 0), pady=6)
                addr = " ".join((o["address"] or o["comment"] or "—").split())
                eta = route["eta"].get(o["id"])
                w = route["windows"][o["id"]] if o["id"] in route["windows"] else point_window(o, day, gcache)
                late = bool(eta is not None and w and w[1] and eta > w[1])
                when = ("~" + fmt_hm(eta)) if eta is not None else ""
                if w:
                    when += ("\n" if when else "") + ("с %s " % fmt_hm(w[0]) if w[0] else "") + \
                        ("до %s" % fmt_hm(w[1]) if w[1] else "")
                vals = [str(r), o["number"], o["agent"], addr, when, money(o["sum"])]
                for c, (val, wd) in enumerate(zip(vals, widths[1:]), 1):
                    wrap_it = c in (3, 4)
                    color = C["muted"] if c == 1 else (C["err"] if (c == 5 and late) else C["text"])
                    if c == 5 and not val:
                        val, color = "часы?", "#B8C1CC"
                    lab = ctk.CTkLabel(line, text=val, width=wd, anchor="e" if c == 6 else "w",
                                       font=F(13, c == 1 or (c == 5 and late)), justify="left",
                                       wraplength=wd - 8 if wrap_it else 0, text_color=color)
                    lab.grid(row=0, column=c, padx=4, pady=4, sticky="w" if wrap_it else "")
                    if c == 5:  # клик по «Приезд» — указать часы приёма
                        lab.configure(cursor="hand2")
                        lab.bind("<Button-1>", lambda e, o=o: edit_hours(o))
                tail = ctk.CTkFrame(line, fg_color="transparent")
                tail.grid(row=0, column=7, padx=(4, 6))
                if o["cash"]:
                    ctk.CTkLabel(tail, text=" нал ", font=F(11, True), fg_color=C["warn_soft"], text_color=C["warn"],
                                 corner_radius=8, height=22).pack(side="left", padx=(0, 4))
                ctk.CTkButton(tail, text="▲", command=lambda o=o: move(o, -1), **arrow).pack(side="left")
                ctk.CTkButton(tail, text="▼", command=lambda o=o: move(o, 1), **arrow).pack(side="left")
            recount()
        render()

        # --- маршрут
        rc = ctk.CTkFrame(box, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["line"])
        rc.pack(fill="x", pady=(12, 0))
        ctk.CTkLabel(rc, text="Маршрут", font=F(12, True), text_color=C["muted"]).grid(
            row=0, column=0, sticky="w", padx=(16, 10), pady=12)
        ctk.CTkLabel(rc, text="выезд", font=F(12), text_color=C["muted"]).grid(row=0, column=1, padx=(0, 4))
        start_var = ctk.StringVar(value=cfg["route_start"])
        ctk.CTkEntry(rc, textvariable=start_var, width=64, height=32, justify="center", font=F(13),
                     border_color=C["line"], fg_color=C["bg"]).grid(row=0, column=2)
        ctk.CTkLabel(rc, text="на точку, мин", font=F(12), text_color=C["muted"]).grid(row=0, column=3, padx=(12, 4))
        svc_var = ctk.StringVar(value=str(cfg["route_service"]))
        ctk.CTkEntry(rc, textvariable=svc_var, width=48, height=32, justify="center", font=F(13),
                     border_color=C["line"], fg_color=C["bg"]).grid(row=0, column=4)
        back_var = ctk.BooleanVar(value=cfg["route_back"])
        ctk.CTkCheckBox(rc, text="вернуться в цех", variable=back_var, font=F(12), checkbox_width=18,
                        checkbox_height=18, fg_color=C["accent"], hover_color=C["accent_hover"]).grid(row=0, column=5, padx=12)
        route_btn = ctk.CTkButton(rc, text="Упорядочить по маршруту", width=210, height=34, corner_radius=10,
                                  font=F(13, True), fg_color=C["accent"], hover_color=C["accent_hover"])
        route_btn.grid(row=0, column=6, padx=(4, 6))
        map_btn = ctk.CTkButton(rc, text="Карта", width=70, height=34, corner_radius=10, font=F(13),
                                fg_color=C["card"], hover_color=C["line"], text_color=C["text"],
                                border_width=1, border_color=C["line"], state="disabled")
        map_btn.grid(row=0, column=7, padx=(0, 6))
        copy_btn = ctk.CTkButton(rc, text="Скопировать ссылку", width=150, height=34, corner_radius=10, font=F(13),
                                 fg_color=C["card"], hover_color=C["line"], text_color=C["text"],
                                 border_width=1, border_color=C["line"], state="disabled")
        copy_btn.grid(row=0, column=8, padx=(0, 12))
        route_info = ctk.CTkLabel(rc, text="Порядок посчитается по дорогам от цеха с учётом «до 16:00» и т. п. из комментариев. "
                                           "Клик по колонке «Приезд» — указать часы работы точки (запомнится).", font=F(12), text_color=C["muted"], justify="left",
                                  anchor="w", wraplength=980)
        route_info.grid(row=1, column=0, columnspan=9, sticky="w", padx=16, pady=(0, 10))
        manual = {}

        def open_map():
            url = current_url()
            if url:
                import webbrowser
                webbrowser.open(url)

        def copy_link():
            url = current_url()
            if url:
                win.clipboard_clear()
                win.clipboard_append(url)
                win.update()
                copy_btn.configure(text="Скопировано ✓")
                win.after(2000, lambda: copy_btn.winfo_exists() and copy_btn.configure(text="Скопировать ссылку"))
        map_btn.configure(command=open_map)
        copy_btn.configure(command=copy_link)

        def ask_missing(missing):
            """Окно: поправить адреса, которых не нашлось на карте."""
            dlg = ctk.CTkToplevel(win, fg_color=C["bg"])
            dlg.title("Не нашёл на карте")
            dlg.transient(win)
            fb = ctk.CTkFrame(dlg, fg_color="transparent")
            fb.pack(padx=20, pady=16)
            ctk.CTkLabel(fb, text="Эти адреса не нашлись на карте", font=F(16, True), text_color=C["text"]).pack(anchor="w")
            ctk.CTkLabel(fb, justify="left", font=F(12), text_color=C["muted"], text=(
                "Впиши адрес попроще: «проспект Ленина 154а» (город не нужен, если это Барнаул).\n"
                "Надёжнее всего — координаты: в Яндекс Картах правый клик по точке → «Что здесь?» →\n"
                "скопировать числа вида «53.34, 83.77». Можно и просто вставить ссылку на точку. Программа запомнит.")).pack(anchor="w", pady=(2, 10))
            vars_ = {}
            for o in missing:
                rr = ctk.CTkFrame(fb, fg_color=C["card"], corner_radius=8)
                rr.pack(fill="x", pady=2)
                tried = manual.get(geo_key(o))
                txt = "%s\n%s" % (o["agent"], " ".join((o["address"] or "нет адреса").split()))
                if tried:
                    txt += "\n⚠ «%s» тоже не нашлось — вставь координаты или ссылку из Яндекс Карт" % tried
                ctk.CTkLabel(rr, text=txt, font=F(12), justify="left", anchor="w", width=360, wraplength=350,
                             text_color=C["warn"] if tried else C["text"]).pack(side="left", padx=10, pady=6)
                v = ctk.StringVar(value=tried or "")
                ctk.CTkEntry(rr, textvariable=v, width=300, height=32, font=F(13), border_color=C["line"],
                             fg_color=C["bg"]).pack(side="left", padx=(0, 10))
                vars_[geo_key(o)] = v

            def again():
                for k, v in vars_.items():
                    if v.get().strip():
                        manual[k] = v.get().strip()
                dlg.destroy()
                do_route()
            bb = ctk.CTkFrame(fb, fg_color="transparent")
            bb.pack(fill="x", pady=(10, 0))
            ctk.CTkButton(bb, text="Пересчитать", width=150, height=38, corner_radius=10, font=F(14, True),
                          fg_color=C["accent"], hover_color=C["accent_hover"], command=again).pack(side="right")
            ctk.CTkButton(bb, text="Пропустить", width=110, height=38, corner_radius=10, font=F(13),
                          fg_color=C["card"], hover_color=C["line"], text_color=C["text"], border_width=1,
                          border_color=C["line"], command=dlg.destroy).pack(side="right", padx=8)
            dlg.after(150, lambda: (dlg.lift(), dlg.focus_force()))

        def do_route():
            import re
            m = re.match(r"^\s*(\d{1,2})\D?(\d{2})?\s*$", start_var.get())
            hm = (int(m.group(1)), int(m.group(2) or 0)) if m else (9, 0)
            try:
                svc = max(0, int(svc_var.get()))
            except ValueError:
                svc = 10
            chosen = [o for o in orders if keep[o["id"]].get()]
            if not chosen:
                route_info.configure(text="Отметь хотя бы одну отгрузку.", text_color=C["err"])
                return
            route_btn.configure(state="disabled", text="Считаю…")
            say_r = lambda t: win.after(0, lambda: route_info.configure(text=t, text_color=C["muted"]))

            def work():
                try:
                    res = plan_route(cfg, day, chosen, hm, svc, back_var.get(), say_r, manual)
                    win.after(0, lambda: applied(res, hm, svc))
                except Oops as e:
                    err = str(e)
                    win.after(0, lambda: failed(err))
                except Exception as e:
                    log_error("маршрут")
                    err = "Ошибка: %s" % e
                    win.after(0, lambda: failed(err))
            threading.Thread(target=work, daemon=True).start()

        def failed(msg):
            route_btn.configure(state="normal", text="Упорядочить по маршруту")
            route_info.configure(text=msg, text_color=C["err"])

        def applied(res, hm, svc):
            route_btn.configure(state="normal", text="Упорядочить по маршруту")
            save_setting("route", "start", "%02d:%02d" % hm)
            save_setting("route", "minutes_per_stop", str(svc))
            save_setting("route", "return_to_depot", "да" if back_var.get() else "нет")
            rest = [o for o in orders if o not in res["order"]]
            orders[:] = res["order"] + rest
            route["eta"], route["windows"] = res["eta"], res["windows"]
            route["points"], route["depot"] = res["points"], res["depot"]
            gcache.clear()
            gcache.update(geo_load())
            has = "normal" if res["points"] else "disabled"
            map_btn.configure(state=has)
            copy_btn.configure(state=has)
            render()
            lates = [o for o in res["order"] if o["id"] in res["eta"] and res["windows"].get(o["id"])
                     and res["windows"][o["id"]][1] and res["eta"][o["id"]] > res["windows"][o["id"]][1]]
            txt = "≈ %.0f км, в пути с остановками ~%d ч %02d мин %s" % (
                res["km"], res["minutes"] // 60, res["minutes"] % 60,
                "· по дорогам" if res["roads"] else "· сервис дорог не ответил — посчитал по прямой, порядок примерный")
            color = C["ok"]
            if lates:
                txt += "\n⚠ Не успеваем: " + ", ".join("%s (до %s, приедем ~%s)" % (
                    o["agent"][:30], fmt_hm(res["windows"][o["id"]][1]), fmt_hm(res["eta"][o["id"]])) for o in lates)
                color = C["warn"]
            if res["missing"]:
                txt += "\nНе нашёл на карте: %d — поставил в конец списка." % len(res["missing"])
                color = C["warn"]
            route_info.configure(text=txt, text_color=color)
            if res["missing"]:
                ask_missing(res["missing"])
        route_btn.configure(command=do_route)

        # тип и название листа
        name_card = ctk.CTkFrame(box, fg_color=C["card"], corner_radius=12, border_width=1, border_color=C["line"])
        name_card.pack(fill="x", pady=(12, 0))
        ctk.CTkLabel(name_card, text="Лист", font=F(12, True), text_color=C["muted"]).grid(
            row=0, column=0, sticky="w", padx=(16, 10), pady=12)
        kind = ctk.CTkSegmentedButton(name_card, values=["Доставка", "Командировка"], height=34, font=F(13),
                                      selected_color=C["accent"], selected_hover_color=C["accent_hover"],
                                      unselected_color=C["bg"], unselected_hover_color=C["line"], text_color=C["text"])
        kind.grid(row=0, column=1, padx=(0, 12))
        ctk.CTkLabel(name_card, text="уточнение", font=F(12), text_color=C["muted"]).grid(row=0, column=2, padx=(0, 6))
        extra_var = ctk.StringVar(value="")
        ctk.CTkEntry(name_card, textvariable=extra_var, width=200, height=34, font=F(13), corner_radius=8,
                     border_color=C["line"], fg_color=C["bg"],
                     placeholder_text="необязательно").grid(row=0, column=3)
        name_var = ctk.StringVar()
        ctk.CTkLabel(name_card, textvariable=name_var, font=F(14, True), text_color=C["text"]).grid(
            row=0, column=4, sticky="w", padx=(14, 6))
        name_hint = ctk.CTkLabel(name_card, text="", font=F(12), text_color=C["warn"])
        name_hint.grid(row=1, column=1, columnspan=4, sticky="w", padx=0, pady=(0, 8))
        kind_touched = {"v": False}

        def check_name(*_):
            name_var.set("→ " + default_name(day, kind.get() or "Доставка", extra_var.get()) + ".xlsx")
            nm = default_name(day, kind.get() or "Доставка", extra_var.get())
            path = free_path(cfg, nm, day)
            name_hint.configure(text="" if os.path.basename(path) == safe_name(nm) + ".xlsx" else
                                "такой уже есть → сохраню как «%s» (+ загрузка с тем же названием)" %
                                os.path.basename(path)[:-5])

        def auto_kind(*_):  # пока не трогали переключатель — подбираем сами по отмеченным отгрузкам
            if not kind_touched["v"]:
                kind.set("Командировка" if looks_like_trip([o for o in orders if keep[o["id"]].get()], cfg)
                         else "Доставка")
                check_name()

        def on_kind(_value):
            kind_touched["v"] = True
            check_name()
        kind.configure(command=on_kind)
        extra_var.trace_add("write", check_name)
        for v in keep.values():
            v.trace_add("write", auto_kind)
        auto_kind()
        check_name()

        foot = ctk.CTkFrame(box, fg_color="transparent")
        foot.pack(fill="x", pady=(14, 0))
        ctk.CTkLabel(foot, textvariable=counter, font=F(14, True), text_color=C["text"]).pack(side="left")

        def set_all(val):
            for v in keep.values():
                v.set(val)

        def go():
            chosen = [o for o in orders if keep[o["id"]].get()]
            name = default_name(day, kind.get() or "Доставка", extra_var.get())
            win.destroy()
            url = current_url()
            background(lambda: save_sheet(chosen, cfg, day, name, route_url=url), sheet_done)
            start_prep(cfg, api, day, orders, {o["id"] for o in chosen})

        small = dict(height=40, corner_radius=10, font=F(13), fg_color=C["card"], hover_color=C["line"],
                     text_color=C["text"], border_width=1, border_color=C["line"])
        ctk.CTkButton(foot, text="Сформировать", width=170, height=44, corner_radius=12, font=F(15, True),
                      fg_color=C["accent"], hover_color=C["accent_hover"], command=go).pack(side="right")
        ctk.CTkButton(foot, text="Отмена", width=90, command=win.destroy, **small).pack(side="right", padx=8)
        ctk.CTkButton(foot, text="Снять все", width=100, command=lambda: set_all(False), **small).pack(side="right", padx=4)
        ctk.CTkButton(foot, text="Выбрать все", width=100, command=lambda: set_all(True), **small).pack(side="right", padx=4)

    def printing():
        day = parse_day(date_var.get())
        if prep.get("day") == day.date() and prep.get("pf"):
            def job_ready():
                prep["ready"].wait()
                if prep.get("error") or prep.get("orders") is None:
                    raise Oops("Не удалось подготовить документы: %s" % prep.get("error"))
                cfg_ = load_config()
                return cfg_, prep["api"], day, sort_like_sheet(cfg_, day, prep["orders"]), prep["pf"]
            background(job_ready, lambda res: choose(*res))
            return

        def job():
            cfg, api, day = ctx()
            return cfg, api, day, plan_print(api, cfg, day, say), None
        background(job, lambda res: choose(*res))

    # --- окно выбора документов
    def choose(cfg, api, day, orders, shared_pf=None):
        busy(False)
        bar.stop()
        stat.pack_forget()
        win = ctk.CTkToplevel(root, fg_color=C["bg"])
        win.title("Что печатаем — %s" % day.strftime("%d.%m.%Y"))
        win.transient(root)
        win.after(100, win.grab_set)
        box = ctk.CTkFrame(win, fg_color="transparent")
        box.pack(fill="both", expand=True, padx=22, pady=18)

        ctk.CTkLabel(box, text="Документы на печать · %s" % day.strftime("%d.%m.%Y"),
                     font=F(20, True), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(box, text="УПД ×%d, по наличке (ГК) — расходная ×%d, ТТН ×%d, счёт ×%d. Сними галочки с того, что не нужно." % (
            cfg["demand_copies"], cfg["cash_copies"], cfg["ttn_copies"], cfg["invoice_copies"]), font=F(12), text_color=C["muted"]).pack(anchor="w", pady=(0, 12))

        widths = (34, 64, 280, 110, 205, 250)
        head = ctk.CTkFrame(box, fg_color="transparent")
        head.pack(fill="x", padx=(10, 26))
        for i, (h, w) in enumerate(zip(["№", "Номер", "Контрагент", "Сумма", "Накладная", "Счёт"], widths)):
            ctk.CTkLabel(head, text=h, width=w, anchor="e" if i == 3 else "w", font=F(12, True),
                         text_color=C["muted"]).grid(row=0, column=i, padx=4)

        lst = ctk.CTkScrollableFrame(box, width=sum(widths) + 60, height=min(46 * len(orders), 470),
                                     fg_color=C["card"], corner_radius=14, border_width=1, border_color=C["line"])
        lst.pack(fill="both", expand=True, pady=(4, 0))

        v_upd, v_inv, v_ttn = [], [], []
        counter = ctk.StringVar()

        def recount(*_):
            u = sum(v.get() for v in v_upd)
            i = sum(v.get() for v in v_inv)
            new = sum(1 for o, v in zip(orders, v_inv) if v.get() and o["inv"]["kind"] == "new")
            rc = sum(1 for o, v in zip(orders, v_upd) if v.get() and o["cash"])
            t = sum(v.get() for v in v_ttn)
            counter.set("К печати:  УПД — %d,  расходных — %d,%s  счетов — %d%s" % (
                u - rc, rc, "  ТТН — %d," % t if t else "", i, "  (новых: %d)" % new if new else ""))

        pill = {"linked": (C["ok_soft"], C["ok"]), "similar": (C["warn_soft"], C["warn"]), "new": (C["bg"], C["muted"])}
        for r, o in enumerate(orders, 1):
            line = ctk.CTkFrame(lst, fg_color=C["row_alt"] if r % 2 == 0 else C["card"], corner_radius=8, height=42)
            line.pack(fill="x", pady=1)
            name = o["agent"] if len(o["agent"]) <= 38 else o["agent"][:37].rstrip() + "…"
            vals = [str(r), o["number"], name, money(o["sum"])]
            for c, (v, w) in enumerate(zip(vals, widths)):
                ctk.CTkLabel(line, text=v, width=w, anchor="e" if c == 3 else "w",
                             font=F(13, c == 0), text_color=C["muted"] if c == 0 else C["text"]).grid(
                    row=0, column=c, padx=4, pady=6)
            vu, vi = ctk.BooleanVar(value=o["print_upd"]), ctk.BooleanVar(value=o["print_inv"])
            vu.trace_add("write", recount)
            vi.trace_add("write", recount)
            v_upd.append(vu)
            v_inv.append(vi)
            vt = ctk.BooleanVar(value=o.get("print_ttn", False))
            vt.trace_add("write", recount)
            v_ttn.append(vt)
            label, _, copies = doc_for(o, cfg)
            cell_doc = ctk.CTkFrame(line, fg_color="transparent", width=widths[4])
            cell_doc.grid(row=0, column=4, padx=4, sticky="w")
            cb = dict(corner_radius=6, checkbox_width=20, checkbox_height=20, fg_color=C["accent"],
                      hover_color=C["accent_hover"], border_color="#B8C1CC")
            ctk.CTkCheckBox(cell_doc, text="%s ×%d" % (label, copies), variable=vu, width=138, font=F(13),
                            **cb).pack(side="left")
            # ТТН — маленькая галочка рядом; у контрагентов из ttn_for стоит сразу
            ctk.CTkCheckBox(cell_doc, text="ТТН", variable=vt, width=60, font=F(12, o["ttn"]),
                            text_color=C["text"] if o["ttn"] else C["muted"], **cb).pack(side="left", padx=(4, 0))
            cell_inv = ctk.CTkFrame(line, fg_color="transparent", width=widths[5])
            cell_inv.grid(row=0, column=5, padx=4, sticky="w")
            ctk.CTkCheckBox(cell_inv, text="", variable=vi, width=24, corner_radius=6, checkbox_width=20,
                            checkbox_height=20, fg_color=C["accent"], hover_color=C["accent_hover"],
                            border_color="#B8C1CC").pack(side="left")
            bg, fg = pill[o["inv"]["kind"]]
            ctk.CTkLabel(cell_inv, text=" %s " % inv_label(o), font=F(12), fg_color=bg, text_color=fg,
                         corner_radius=8, height=26).pack(side="left", padx=(2, 0))
        recount()

        # документы готовятся заранее (или уже готовятся — после маршрутного листа)
        pf = shared_pf or Prefetcher(api)
        pf.warm(orders, cfg)

        def on_toggle(o, vu, vi, vt):  # поставили галочку — сразу начинаем качать
            if vu.get():
                _, name, copies = doc_for(o, cfg)
                pf.get("demand", o["id"], name, copies)
            if vt.get():
                pf.get("demand", o["id"], cfg["ttn_template"], cfg["ttn_copies"])
            if vi.get() and o["inv"]["kind"] != "new":
                pf.get("invoiceout", o["inv"]["id"], cfg["invoice_template"], cfg["invoice_copies"])
        for o, vu, vi, vt in zip(orders, v_upd, v_inv, v_ttn):
            for var in (vu, vi, vt):
                var.trace_add("write", lambda *_, o=o, vu=vu, vi=vi, vt=vt: on_toggle(o, vu, vi, vt))
        ready = ctk.StringVar()

        def tick():
            if not win.winfo_exists():
                return
            d, n = pf.progress()
            ready.set("Готовлю документы заранее: %d из %d" % (d, n) if d < n else "Документы готовы ✓")
            win.after(500, tick)
        tick()

        foot = ctk.CTkFrame(box, fg_color="transparent")
        foot.pack(fill="x", pady=(14, 0))
        left = ctk.CTkFrame(foot, fg_color="transparent")
        left.pack(side="left")
        ctk.CTkLabel(left, textvariable=counter, font=F(14, True), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(left, textvariable=ready, font=F(12), text_color=C["muted"]).pack(anchor="w")

        def cancel():
            if not shared_pf:
                pf.close()
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", cancel)

        def go():
            for o, vu, vi, vt in zip(orders, v_upd, v_inv, v_ttn):
                o["print_upd"], o["print_inv"], o["print_ttn"] = vu.get(), vi.get(), vt.get()
            win.destroy()

            if shared_pf:
                prep.clear()  # набор передаём печати; закроется после неё

            stop = threading.Event()

            def job():
                try:
                    return do_print(api, cfg, day, orders, say, pf, cancel=stop)
                finally:
                    pf.close()
            background(job, lambda res: finish(
                True, "✓  Готово: УПД — %d, расходных — %d,%s счетов — %d%s" % (
                    res[1], res[4], " ТТН — %d," % res[5] if res[5] else "", res[2],
                    ("\nСоздано новых счетов: %d" % res[3] if res[3] else "") +
                    ("\n\n⚠ %s. Подробности — в файле «ошибки.log»." % res[6] if res[6] else "")),
                res[0], lambda p: send_to_print(p, cfg["print_mode"])), cancel=stop)

        def set_all(val):
            for v in v_upd:
                v.set(val)

        small = dict(height=40, corner_radius=10, font=F(13), fg_color=C["card"], hover_color=C["line"],
                     text_color=C["text"], border_width=1, border_color=C["line"])
        ctk.CTkButton(foot, text="Печать", width=150, height=44, corner_radius=12, font=F(15, True),
                      fg_color=C["accent"], hover_color=C["accent_hover"], command=go).pack(side="right")
        ctk.CTkButton(foot, text="Отмена", width=90, command=cancel, **small).pack(side="right", padx=8)
        ctk.CTkButton(foot, text="Снять все", width=100, command=lambda: set_all(False), **small).pack(side="right", padx=4)
        ctk.CTkButton(foot, text="Выбрать все", width=100, command=lambda: set_all(True), **small).pack(side="right", padx=4)

    def analysis():
        def run(months, paths, min_orders):
            def job():
                cfg = load_config()
                return analyze_clients(make_api(cfg), cfg, months, paths, min_orders, say)
            background(job, lambda res: finish(True, (
                "✓  Отчёт готов: клиентов %(клиентов)d\n"
                "Постоянных, но не в списках: %(не в списках)d\n"
                "Пора звонить: %(пора звонить)d\n"
                "Строк списков без заказов: %(без заказов)d") % res[1], res[0]))
        analysis_dialog(root, F, run)

    btn.configure(command=sheet)
    btn2.configure(command=printing)
    btn3.configure(command=analysis)
    root.bind("<Return>", lambda e: sheet())
    if not has_token():
        root.after(300, lambda: settings_dialog(root, F, first_run=True))

    def tidy():  # старые файлы из общей папки — по папкам дней
        try:
            tidy_old_files(load_config())
        except Exception:
            pass
    threading.Thread(target=tidy, daemon=True).start()
    root.mainloop()


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        if check_update():
            return restart()
        return gui()
    try:
        cfg = load_config()
        api = make_api(cfg)
        day = parse_day(args[0])
        if "--dump" in sys.argv:
            print("Сохранено: %s" % dump_fields(api, day))
            return
        if "--print" in sys.argv:  # python marshrut.py 22.09.2026 --print  (всё по умолчанию)
            orders = plan_print(api, cfg, day, print)
            for o in orders:
                print("%-6s %-40s %-10s %s" % (o["number"], o["agent"][:40], doc_for(o, cfg)[0],
                                                 inv_label(o) if o["print_inv"] else ""))
            path, u, i, new, rc, _ttn, note = do_print(api, cfg, day, orders, print)
            if note:
                print("Внимание:", note)
            print("Готово: %s\nУПД: %d, расходных: %d, счетов: %d, создано новых: %d" % (path, u, rc, i, new))
            if "--no-open" not in sys.argv:
                send_to_print(path, cfg["print_mode"])
            return
        path, n, cash, hint, total, _load = generate(api, cfg, day)
    except Oops as e:
        sys.exit(str(e))
    print("Готово: %s\nДоставок: %d на %.2f руб., наличка: %.2f руб. %s" % (path, n, total, cash, hint))
    if "--no-open" not in sys.argv:
        open_file(path)


if __name__ == "__main__":
    main()
