import os, io, re, json, base64, calendar, datetime, shutil, tempfile, threading, webbrowser, traceback
import importlib.util
from PIL import Image
import win32com.client as win32
import anthropic
from dash import Dash, dcc, html, dash_table, Input, Output, State, no_update, ctx

# ---------- Konfiguration ----------
from dotenv import load_dotenv; load_dotenv(r"S:\benjaminSuermann\3_env\.env")
API_KEY = os.environ["ANTHROPIC_API_KEY"]
MODEL = "claude-opus-4-8"
TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bviSheetOutline.xls")  # Stilvorlage
OUTDIR = r"Q:\7_NTP_nordIX_Treasury_plus\1_NAD_Manager\1_bvi"
SHEET = "BVI_Securities"
FIRST_ROW = 11
PORT = 8055
MAXEDGE = 2400
FORCE_OFFSET = None   # "+01:00"/"+02:00" erzwingen, sonst automatisch (Europe/Berlin)
THEME_PATH = r"S:\benjaminSuermann\3_env\pyDashDesign.py"

# Fallback, falls THEME_PATH fehlt/fehlerhaft (Werte entsprechen pythonDesigns.py)
_THEME_C = {"background": "#FAF8F5", "surface": "#EDEBE8", "border": "#C9C6C1",
            "primary": "#5C7285", "secondary": "#B99A7C", "text": "#2A2A28",
            "positive": "#9FC1CE", "negative": "#C98CA7", "highlight": "#D9A76A"}
_THEME_F = {"family": "Georgia, 'Times New Roman', serif", "size_title": 28,
            "size_subtitle": 18, "size_body": 14, "size_label": 11,
            "weight_title": 300, "weight_body": 400}


def load_design():
    """Laedt die Designdatei (THEME_PATH) frisch. Gibt das Modul zurueck oder None."""
    try:
        spec = importlib.util.spec_from_file_location("_bvi_design", THEME_PATH)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    except Exception:
        return None


def theme():
    """Liest Farben/Fonts/Logo bei jedem Aufruf frisch aus THEME_PATH (Aenderungen -> Browser-Refresh)."""
    C, F, logo_src = dict(_THEME_C), dict(_THEME_F), None
    mod = load_design()
    if mod is not None:
        if isinstance(getattr(mod, "COLORS", None), dict):
            C.update(mod.COLORS)
        if isinstance(getattr(mod, "FONT", None), dict):
            F.update(mod.FONT)
        logo_src = getattr(mod, "LOGO_SRC", None)
    return C, F, logo_src

# Bloomberg-Account -> Portfolio (BVI Spalten D/E/F)
PORTFOLIOS = {
    "42005137": {"D": "082L00", "E": "082L01", "F": "nordIX Anleihen Defensiv"},
    "61212723": {"D": "082L00", "E": "082L01", "F": "nordIX Anleihen Defensiv"},
}
DEFAULT_ACCOUNT = "42005137"

# Broker-Name (Bloomberg) -> (BIC, Langname)
COUNTERPARTIES = [
    (("BARCLAYS",),                    "BARCIE2D",    "Barclays Bank Ireland PLC"),
    (("JP MORGAN", "JPMORGAN", "JPM"), "CHASDEFXXXX", "J.P. Morgan AG"),
    (("GOLDMAN", "GSA"),               "GOLDDEFAXXX", "Goldman Sachs Bank Europe SE"),
    (("UBS", "EUBS"),                  "UBSWDE24XXX", "UBS Europe SE"),
    (("DEUTSCHE",),                    "DEUTDEFFDSO", "Deutsche Bank AG"),
    (("HSBC",),                        "TUBDDEDDXXX", "HSBC (D)"),
    (("DONNER", "REUSCHEL"),           "CHDBDEHHXXX", "Donner & Reuschel AG"),
]

COLS = [("side", "Buy/Sell"), ("isin", "ISIN"), ("name", "Bezeichnung"), ("qty", "Menge"),
        ("price", "Clean Price"), ("ccy", "CCY"), ("interest", "Stueckzinsen"), ("int_days", "Zinstage"),
        ("settle_amt", "Net"), ("trade_date", "Trade Date"), ("exec_time", "Exec Time"),
        ("settle_date", "Settle Date"), ("account", "Account"), ("broker_name", "Broker Name"),
        ("broker_bic", "Broker BIC"), ("pf_kvg", "Portfolio KVG")]
FIELDS = [c[0] for c in COLS]

# ---------- Parser / Helfer ----------

def num(s):
    if s is None or s == "":
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace("€", "").replace("$", "").replace(" ", "").replace("\xa0", "")
    if s == "":
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "")
    return float(s)


def to_date(s):
    if isinstance(s, datetime.date):
        return s
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%m/%d/%Y", "%d/%m/%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Datum nicht erkannt: {s!r}")


def last_sunday(y, m):
    for week in reversed(calendar.monthcalendar(y, m)):
        if week[calendar.SUNDAY]:
            return datetime.date(y, m, week[calendar.SUNDAY])


def offset_for(d):
    if FORCE_OFFSET:
        return FORCE_OFFSET
    return "+02:00" if last_sunday(d.year, 3) <= d < last_sunday(d.year, 10) else "+01:00"


def exec_timestamp(trade_date, time_str):
    time_str = str(time_str).strip()
    if not time_str:
        return ""
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.datetime.strptime(time_str, fmt).time()
            return f"{trade_date.isoformat()}T{t.strftime('%H:%M:%S')}{offset_for(trade_date)}"
        except ValueError:
            continue
    raise ValueError(f"Uhrzeit nicht erkannt: {time_str!r}")


def map_side(s):
    s = str(s).strip().lstrip("﻿").upper()
    if s in ("S", "SELL", "SE", "VERKAUF", "V"):
        return "SELL"
    if s in ("B", "BUY", "BUYI", "BY", "KAUF", "K"):
        return "BUYI"
    return s


def resolve_broker(name):
    if not name:
        return None
    u = str(name).upper()
    for keys, bic, full in COUNTERPARTIES:
        if any(k in u for k in keys):
            return bic, full
    return None


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def try_date(s):
    if s in (None, ""):
        return None
    try:
        d = to_date(s)
    except Exception:
        return None
    return d if 2000 <= d.year <= 2100 else None


def try_num(s):
    try:
        return num(s)
    except Exception:
        return None


def ticker_of(name):
    tok = re.split(r"\s+", str(name or "").strip())
    t = re.sub(r"[^A-Za-z0-9]", "", tok[0]) if tok and tok[0] else ""
    return t.upper() or "NA"

# ---------- Excel: Stilvorlage kopieren und Datenzeilen fuellen ----------

def write_workbook(dest, rows):
    if not os.path.exists(TEMPLATE):
        raise RuntimeError(f"Stilvorlage nicht gefunden: {TEMPLATE}")
    tmp = os.path.join(tempfile.gettempdir(), f"_bvi_tpl_{os.getpid()}_{abs(id(rows))}.xls")
    try:
        shutil.copy2(TEMPLATE, tmp)
    except Exception as e:
        raise RuntimeError(f"Vorlage nicht lesbar ({e}). Noch in Excel geoeffnet?")
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(tmp, IgnoreReadOnlyRecommended=True)
        ws = wb.Worksheets(SHEET)
        used = ws.UsedRange
        last = used.Row + used.Rows.Count - 1
        if last >= FIRST_ROW:                       # Beispiel-/Altzeilen leeren, Format bleibt
            ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(last, 40)).ClearContents()
        for i, row in enumerate(rows):
            r = FIRST_ROW + i
            for col, val in row.items():
                if val == "" or val is None:
                    continue
                c = ws.Cells(r, col_num(col))
                if col in ("U", "W"):                      # Datum, kurz (Zellenformat Datum)
                    c.NumberFormatLocal = "TT.MM.JJJJ"
                    c.Value = (val - datetime.date(1899, 12, 30)).days
                elif col == "V":
                    c.NumberFormatLocal = "@"
                    c.Value = val
                elif col == "L":                           # Clean Price: alle Nachkommastellen zeigen
                    c.NumberFormatLocal = "0,##########"
                    c.Value = val
                else:
                    c.Value = val
        wb.SaveAs(dest, FileFormat=56)   # 56 = xlExcel8 (.xls, wie die Vorlage)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
        try:
            os.remove(tmp)
        except Exception:
            pass


def build_bvi_row(r):
    ccy = str(r.get("ccy") or "EUR").upper()
    td, sd = to_date(r["trade_date"]), to_date(r["settle_date"])
    idays = r.get("int_days")
    return {"A": "", "B": "NEWM", "C": "", "D": "082L00", "E": r.get("pf_kvg") or "082L01",
            "F": "nordIX Anleihen Defensiv", "G": map_side(r["side"]), "H": num(r["qty"]),
            "I": "ISIN", "J": str(r["isin"]).upper().strip(), "K": r.get("name", ""),
            "L": num(r["price"]), "M": ccy, "N": 0.0, "O": 0.0, "P": 0.0, "Q": 0.0,
            "R": num(r.get("interest")) or 0.0, "S": num(r["settle_amt"]),
            "T": int(num(idays)) if str(idays) not in ("None", "", "0") else "",
            "U": td, "V": exec_timestamp(td, str(r.get("exec_time") or "")), "W": sd,
            "X": ccy, "Y": "XOFF", "Z": "BIC", "AA": str(r.get("broker_bic") or "").upper(),
            "AB": r.get("broker_name") or "", "AC": "", "AD": "", "AE": 1.0, "AF": 1.0, "AG": ""}


def validate(rows):
    errs = []
    for i, r in enumerate(rows, 1):
        for f, lab in (("isin", "ISIN"), ("name", "Bezeichnung"), ("side", "Buy/Sell")):
            if not str(r.get(f, "")).strip():
                errs.append(f"Zeile {i}: {lab} fehlt")
        for f, lab in (("qty", "Menge"), ("price", "Clean Price"), ("settle_amt", "Net")):
            if try_num(r.get(f)) is None:
                errs.append(f"Zeile {i}: {lab} ungueltig ('{r.get(f)}')")
        for f, lab in (("trade_date", "Trade Date"), ("settle_date", "Settle Date")):
            if try_date(r.get(f)) is None:
                errs.append(f"Zeile {i}: {lab} kein gueltiges Datum ('{r.get(f)}')")
        if r.get("exec_time"):
            try:
                exec_timestamp(datetime.date(2000, 1, 1), str(r["exec_time"]))
            except Exception:
                errs.append(f"Zeile {i}: Exec Time ungueltig ('{r.get('exec_time')}')")
    return errs


def unique_dest(base):
    dest = os.path.join(OUTDIR, base + ".xls")
    n = 2
    while os.path.exists(dest):
        dest = os.path.join(OUTDIR, f"{base}_{n}.xls")
        n += 1
    return dest

# ---------- Bilderkennung (Claude Vision) ----------
SCHEMA = {"type": "object", "additionalProperties": False,
    "properties": {"trades": {"type": "array", "items": {"type": "object", "additionalProperties": False,
        "properties": {k: ({"type": "number"} if k in ("qty", "price", "interest", "settle_amt")
                           else {"type": "integer"} if k == "int_days" else {"type": "string"})
                       for k in ("side", "isin", "name", "qty", "price", "ccy", "interest", "int_days",
                                 "settle_amt", "trade_date", "exec_time", "settle_date", "account", "broker")},
        "required": ["side", "isin", "name", "qty", "price", "ccy", "interest", "int_days",
                     "settle_amt", "trade_date", "exec_time", "settle_date", "account", "broker"]}}},
    "required": ["trades"]}

PROMPT = """Die Quellen (Screenshots/PDF/Text) enthalten Bloomberg-Wertpapier-Trades (BLOT-Tickets).
Lies ALLE erkennbaren Trades exakt aus. Feld-Zuordnung je Trade:
side="Buy/Sell"; isin="ISIN"; name="Issue"; qty="Quantity"(Zahl); price="Clean Price";
ccy=Waehrung(Euro-Zeichen="EUR"); interest="Acc Int"(Betrag); int_days=Zahl in "Acc Int (NNN)";
settle_amt="Net"; trade_date="Trade Date" als YYYY-MM-DD (Bloomberg zeigt MM/DD/YYYY);
exec_time="Entry/Exec Time" ZWEITE Uhrzeit als HH:MM:SS; settle_date="Settle Date" als YYYY-MM-DD;
account="Account"; broker="Broker Name". Betraege als reine Zahlen (Punkt=Dezimal, keine Tausender).
Datumsangaben IMMER mit vierstelligem Jahr (2026-...); niemals Platzhalter wie "yyyy".
Ist ein Datum nicht sicher lesbar, lass das Feld leer."""


def img_block(raw):
    im = Image.open(io.BytesIO(raw))
    im.load()
    if im.mode != "RGB":
        im = im.convert("RGB")
    m = max(im.size)
    if m > MAXEDGE:
        s = MAXEDGE / m
        im = im.resize((max(1, int(im.width * s)), max(1, int(im.height * s))), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return {"type": "image", "source": {"type": "base64", "media_type": "image/png",
            "data": base64.standard_b64encode(buf.getvalue()).decode("ascii")}}


def build_content(sources):
    blocks, texts = [], []
    for url, fn in sources:
        raw = base64.b64decode(url.split(",", 1)[1])
        ext = os.path.splitext(fn)[1].lower()
        head = url[:30].lower()
        if ext in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff") or "image/" in head:
            blocks.append(img_block(raw))
        elif ext == ".pdf" or "pdf" in head:
            blocks.append({"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                           "data": base64.standard_b64encode(raw).decode("ascii")}})
        else:
            texts.append(f"[{fn}]\n" + raw.decode("utf-8", "replace"))
    if texts:
        blocks.append({"type": "text", "text": "Text-Quellen:\n\n" + "\n\n".join(texts)})
    blocks.append({"type": "text", "text": PROMPT})
    return blocks


def read_trades(sources):
    cl = anthropic.Anthropic(api_key=API_KEY)
    res = cl.messages.create(model=MODEL, max_tokens=8192,
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
        messages=[{"role": "user", "content": build_content(sources)}])
    text = next(b.text for b in res.content if b.type == "text")
    return json.loads(text).get("trades", [])


def to_row(t):
    bic, bname = "", t.get("broker", "")
    r = resolve_broker(t.get("broker", ""))
    if r:
        bic, bname = r
    acct = str(t.get("account") or "")
    pf = PORTFOLIOS.get(acct, PORTFOLIOS[DEFAULT_ACCOUNT])
    return {"side": t.get("side", ""), "isin": t.get("isin", ""), "name": t.get("name", ""),
            "qty": t.get("qty", ""), "price": t.get("price", ""), "ccy": t.get("ccy") or "EUR",
            "interest": t.get("interest", 0), "int_days": t.get("int_days", ""),
            "settle_amt": t.get("settle_amt", ""), "trade_date": t.get("trade_date", ""),
            "exec_time": t.get("exec_time", ""), "settle_date": t.get("settle_date", ""),
            "account": acct, "broker_name": bname, "broker_bic": bic, "pf_kvg": pf["E"]}

# ---------- Dashboard ----------
app = Dash(__name__, title="BVI-Generator")

# Strg+V (Screenshot aus Zwischenablage) -> dcc.Store 'pasted'
app.index_string = """<!DOCTYPE html><html><head>{%metas%}<title>{%title%}</title>{%favicon%}{%css%}</head>
<body>{%app_entry%}<footer>{%config%}{%scripts%}{%renderer%}</footer>
<script>
document.addEventListener('paste', function (e) {
  var t = e.target;
  if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
  var items = (e.clipboardData || window.clipboardData).items; if (!items) return;
  for (var i = 0; i < items.length; i++) {
    if (items[i].type && items[i].type.indexOf('image') === 0) {
      var blob = items[i].getAsFile(); var reader = new FileReader();
      reader.onload = function (ev) {
        if (window.dash_clientside && window.dash_clientside.set_props)
          window.dash_clientside.set_props('pasted', { data: { url: ev.target.result, t: Date.now() } });
      };
      reader.readAsDataURL(blob); e.preventDefault(); return;
    }
  }
});
</script></body></html>"""


def serve_layout():
    C, F, logo_src = theme()             # bei jedem Seiten-Refresh frisch aus pyDashDesign.py
    fam = F.get("family", "Georgia, serif")
    BG, SURF, LINE = C["background"], C["surface"], C["border"]
    NAVY, ACCENT, SEC, TXT = C["primary"], C["primary"], C["secondary"], C["text"]

    def card(children):
        return html.Div(children, style={"background": SURF, "border": f"1px solid {LINE}",
            "borderRadius": "8px", "padding": "18px", "marginBottom": "16px"})

    brand = []
    if logo_src:                         # Logo links oben (eingebettet aus der Designdatei)
        brand.append(html.Img(src=logo_src, style={"height": "46px", "display": "block"}))
    brand.append(html.Div([
        html.Div("BVI-Generator", style={"fontSize": f"{F['size_title']}px",
                 "fontWeight": F["weight_title"], "color": TXT, "lineHeight": "1.1"}),
        html.Div("nordIX Anleihen Defensiv · Universal Investment (082L01)",
                 style={"fontSize": "12px", "textTransform": "uppercase", "letterSpacing": ".5px",
                        "color": SEC, "marginTop": "4px"})]))
    header = html.Div(style={"background": "#FFFFFF", "padding": "14px 26px", "display": "flex",
                             "alignItems": "center", "gap": "16px",
                             "borderBottom": f"3px solid {NAVY}"}, children=brand)

    upload = card([dcc.Upload(id="up", multiple=True, accept="image/*,application/pdf,text/*",
        children=html.Div([
            html.Div("\U0001F4CB  Screenshot einfuegen  (Strg + V)",
                     style={"fontSize": "16px", "fontWeight": "600", "color": NAVY}),
            html.Div("oder Dateien hierher ziehen / klicken  ·  Screenshots, PDF, Text  ·  mehrere moeglich",
                     style={"fontSize": "12.5px", "color": SEC, "marginTop": "6px"})]),
        style={"padding": "26px", "border": f"2px dashed {ACCENT}", "borderRadius": "8px",
               "textAlign": "center", "background": BG, "cursor": "pointer"})])

    action_bar = html.Div(style={"display": "flex", "alignItems": "center", "gap": "16px",
        "flexWrap": "wrap", "marginTop": "14px", "paddingTop": "14px", "borderTop": f"1px solid {LINE}"}, children=[
        html.Button("+ Zeile", id="add", n_clicks=0, style={"padding": "7px 12px", "borderRadius": "6px",
            "border": f"1px solid {LINE}", "background": BG, "color": TXT, "cursor": "pointer", "fontFamily": fam}),
        html.Button("Tabelle leeren", id="clear", n_clicks=0, style={"padding": "7px 12px",
            "borderRadius": "6px", "border": f"1px solid {LINE}", "background": BG, "color": TXT,
            "cursor": "pointer", "fontFamily": fam}),
        html.Div("je Trade eine BVI-Datei  ·  Name: JJJJMMTT_Ticker",
                 style={"fontSize": "12px", "color": SEC}),
        html.Div(style={"flex": "1"}),
        html.Button("BVI erstellen & ablegen", id="save", n_clicks=0, style={"background": NAVY,
            "color": "#fff", "border": "none", "padding": "10px 20px", "borderRadius": "6px",
            "cursor": "pointer", "fontWeight": "700", "fontSize": "14px", "fontFamily": fam})])

    table_card = card([
        html.Div("Trades", style={"fontSize": "15px", "fontWeight": "700", "color": NAVY, "marginBottom": "10px"}),
        dash_table.DataTable(id="tbl", columns=[{"name": n, "id": i} for i, n in COLS], data=[],
            editable=True, row_deletable=True, style_table={"overflowX": "auto"}, style_as_list_view=True,
            style_cell={"fontFamily": fam, "fontSize": "13px", "padding": "7px 8px",
                        "minWidth": "72px", "textAlign": "left", "border": "none", "color": TXT,
                        "backgroundColor": SURF, "borderBottom": f"1px solid {LINE}", "whiteSpace": "normal"},
            style_header={"fontWeight": "700", "color": NAVY, "background": BG,
                          "borderBottom": f"2px solid {LINE}"},
            style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": BG},
                                    {"if": {"column_id": "side"}, "fontWeight": "600"}]),
        action_bar])

    return html.Div(style={"background": BG, "minHeight": "100vh", "fontFamily": fam, "color": TXT}, children=[
        dcc.Store(id="pasted"), header,
        html.Div(style={"maxWidth": "1220px", "margin": "22px auto", "padding": "0 18px"}, children=[
            upload,
            dcc.Loading(type="dot", children=html.Div(id="msg", style={"minHeight": "20px",
                "margin": "0 0 12px 2px", "fontSize": "13px", "color": ACCENT, "fontWeight": "600"})),
            table_card,
            html.Div(id="out", style={"whiteSpace": "pre-wrap", "fontSize": "14px"}),
            html.Div(f"Ablage:  {OUTDIR}", style={"marginTop": "22px", "color": SEC, "fontSize": "12px"})])])


app.layout = serve_layout


@app.callback(Output("tbl", "data"), Output("msg", "children"),
              Input("up", "contents"), Input("pasted", "data"),
              Input("add", "n_clicks"), Input("clear", "n_clicks"),
              State("up", "filenames"), State("tbl", "data"), prevent_initial_call=True)
def on_input(contents, pasted, add_c, clear_c, names, data):
    data = data or []
    trig = ctx.triggered_id
    if trig == "clear":
        return [], ""
    if trig == "add":
        return data + [{f: "" for f in FIELDS}], no_update
    if trig == "pasted":
        if not pasted or not pasted.get("url"):
            return no_update, no_update
        sources = [(pasted["url"], "einfuegen.png")]
    elif trig == "up":
        if not contents:
            return no_update, no_update
        sources = list(zip(contents, names or [f"datei{i}" for i in range(len(contents))]))
    else:
        return no_update, no_update
    try:
        trades = read_trades(sources)
    except Exception as e:
        traceback.print_exc()
        return no_update, f"Fehler beim Auslesen: {e}"
    if not trades:
        return no_update, "Keine Trades erkannt – Bild groesser/deutlicher einfuegen."
    return data + [to_row(t) for t in trades], f"{len(trades)} Trade(s) ausgelesen – bitte pruefen."


def _errbox(title, lines):
    C, _, _ = theme()
    return html.Div(style={"background": C["negative"], "border": f"1px solid {C['border']}",
        "borderRadius": "8px", "padding": "12px 14px"}, children=[
        html.Div(title, style={"fontWeight": "700", "color": C["text"]}),
        *[html.Div(x, style={"fontSize": "12.5px", "color": C["text"]}) for x in lines]])


@app.callback(Output("out", "children"),
              Input("save", "n_clicks"), State("tbl", "data"), prevent_initial_call=True)
def on_save(n, data):
    rows_in = [r for r in (data or []) if str(r.get("isin", "")).strip()]
    if not rows_in:
        return _errbox("Keine Zeilen zum Speichern.", [])
    errs = validate(rows_in)
    if errs:
        return _errbox("Bitte zuerst korrigieren:", errs)
    saved = []
    try:
        for r in rows_in:
            d = try_date(r.get("trade_date"))
            base = f"{d.strftime('%Y%m%d')}_{ticker_of(r.get('name'))}"
            dest = unique_dest(base)
            write_workbook(dest, [build_bvi_row(r)])
            saved.append(dest)
    except Exception as e:
        traceback.print_exc()
        return _errbox(f"Fehler beim Speichern: {e}", [])
    C, _, _ = theme()
    return html.Div(style={"background": C["positive"], "border": f"1px solid {C['border']}",
        "borderRadius": "8px", "padding": "12px 14px"}, children=[
        html.Div(f"✓  {len(saved)} BVI-Datei(en) gespeichert", style={"fontWeight": "700", "color": C["text"]}),
        *[html.Div(os.path.basename(p), style={"fontSize": "12.5px", "color": C["text"]}) for p in saved]])

# ---------- Start ----------

def free_port(port):
    import subprocess
    try:
        raw = subprocess.run(["netstat", "-ano", "-p", "tcp"], capture_output=True).stdout or b""
    except Exception:
        return
    out = raw.decode("utf-8", "ignore")
    mypid, killed = str(os.getpid()), set()
    for ln in out.splitlines():
        p = ln.split()
        if len(p) >= 5 and p[1].endswith(f":{port}") and p[2] in ("0.0.0.0:0", "[::]:0", "*:*"):
            pid = p[-1]
            if pid.isdigit() and pid != mypid and pid not in killed:
                subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True)
                killed.add(pid)


def run():
    try:
        free_port(PORT)
    except Exception:
        pass
    url = f"http://127.0.0.1:{PORT}"
    print("\n  BVI-Dashboard laeuft:  " + url)
    print("  Falls sich nichts oeffnet: obige Adresse im Browser aufrufen.  Beenden: Fenster schliessen.\n")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    try:
        app.run(debug=False, port=PORT)
    except OSError as e:
        print(f"\n  Port {PORT} belegt ({e}). Bitte altes Fenster schliessen und neu starten.\n")
        input("  Enter zum Beenden...")


if __name__ == "__main__":
    run()
