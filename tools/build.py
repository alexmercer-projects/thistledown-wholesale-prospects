#!/usr/bin/env python3
"""Builds the prospect list page, the CSV and the Excel workbook from tools/data.py.

    python3 tools/build.py

Everything is worked out from the rows, so the counts on the page, the scores
and the files always agree.
"""
import csv, html, os, re, sys
from collections import Counter, OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import data as D

DOCS = os.path.join(HERE, "..", "docs")
XLSX = "Thistledown-wholesale-prospects-DEMO.xlsx"
CSVF = "Thistledown-wholesale-prospects-DEMO.csv"
e = html.escape


STOP = {"the", "and", "co", "of", "at"}
_taken = {}


def slug(name):
    """Short handle for the placeholder email and website, e.g. 'copperkettle'."""
    if name in _taken:
        return _taken[name]
    words = [w for w in re.findall(r"[a-z0-9]+", name.lower().split("(")[0]) if w not in STOP]
    for n in (2, 3, 4):
        s = "".join(words[:n])
        if s not in _taken.values():
            break
    _taken[name] = s
    return s


SRC_PATH = {"site": "about", "menu": "menu", "listing": "listing", "ig": "social", "license": "license"}


def dist_pts(miles):
    return 2 if miles <= 12 else 1


def build_rows():
    out = []
    for i, r in enumerate(D.ROWS):
        (name, typ, town, miles, street, contact, role, has_email, now,
         c_coffee, c_vol, c_local, c_open, src, day, why, opener) = r
        s = slug(name)
        d = dist_pts(miles)
        score = c_coffee + c_vol + d + c_local + c_open
        tier = "A" if score >= 8 else "B" if score >= 6 else "C"
        out.append(OrderedDict(
            id="TD-%03d" % (i + 1), name=name, type=typ, town=town, miles=miles,
            address="%s, %s" % (street, town), contact=contact, role=role,
            email=("%s.%s@example.com" % (contact.split()[0].lower(), s)) if has_email else "",
            email_note="Published on their site" if has_email else "Contact form only",
            phone="555-01%02d" % i, website="https://example.com/%s" % s,
            coffee_now=now, c_coffee=c_coffee, c_volume=c_vol, c_distance=d, c_local=c_local, c_open=c_open,
            score=score, tier=tier, why=why, opener=opener,
            source=D.SOURCES[src], source_url="https://example.com/%s/%s" % (s, SRC_PATH[src]),
            checked="2026-11-%s" % day, status="Not contacted",
        ))
    return out


ROWS = build_rows()
N = len(ROWS)
TIERS = Counter(r["tier"] for r in ROWS)
EMAILS = sum(1 for r in ROWS if r["email"])
FORMS = N - EMAILS


# ----------------------------------------------------------------- CSV
def write_csv():
    cols = ["id", "name", "type", "town", "miles", "address", "contact", "role", "email", "email_note", "phone",
            "website", "coffee_now", "score", "tier", "why", "opener", "source", "source_url", "checked", "status"]
    heads = ["ID", "Business", "Type", "Town", "Miles from roastery", "Address", "Contact", "Role", "Email",
             "Email note", "Phone", "Website", "Coffee now", "Fit score (of 10)", "Tier", "Why it fits",
             "Suggested first line", "Source", "Source URL", "Checked", "Status"]
    with open(os.path.join(DOCS, CSVF), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(heads)
        for r in ROWS:
            w.writerow([r[c] for c in cols])


# ----------------------------------------------------------------- XLSX
def write_xlsx():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter as L

    INK, ROAST, PAPER = "1A1714", "8A3C1C", "F4F0E8"
    head_font = Font(name="Calibri", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor=ROAST)
    thin = Side(style="thin", color="D9D2C5")
    wrap = Alignment(wrap_text=True, vertical="top")

    wb = Workbook()
    # Read Me
    ws = wb.active
    ws.title = "Read Me"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 92
    ws["B2"] = "Thistledown Coffee Roasters: wholesale prospect list"
    ws["B2"].font = Font(size=18, bold=True, color=INK)
    ws["B3"] = ("CONCEPT DEMO by Alex Mercer. Thistledown, Harlow Valley and every business, person, address, "
                "phone number (555-01xx) and email (example.com) in this file are made up. Example data, not client work.")
    ws["B3"].font = Font(italic=True, color="B4361F")
    ws.merge_cells("B3:C3")
    ws.row_dimensions[3].height = 32
    ws["B3"].alignment = wrap
    info = [
        ("The ask", "Independent places inside the 25-mile delivery loop that could buy coffee wholesale, with who to contact, why each fits, and where the information came from."),
        ("Prospects", "%d rows on the Prospects tab. Tier A: %d, Tier B: %d, Tier C: %d." % (N, TIERS["A"], TIERS["B"], TIERS["C"])),
        ("Fit score", "Five parts added up by formula (see the Scoring tab). Change a part and the score and tier update."),
        ("Emails", "%d emails, each published on the business's own site. %d rows say 'Contact form only'. No emails were guessed." % (EMAILS, FORMS)),
        ("Sources", "Every row has a source link and the date it was checked (%s)." % D.CHECKED_RANGE),
        ("Left off", "%d places found but left off, each with the rule it failed (Skipped tab). %d duplicates merged (Duplicates tab)." % (len(D.SKIPPED), len(D.DUPES))),
        ("Status column", "Pick from the dropdown as you work the list: " + ", ".join(D.STATUS) + "."),
        ("Summary tab", "Counts by type, town and tier, worked out with COUNTIFS from the Prospects tab."),
    ]
    for i, (k, v) in enumerate(info):
        ws.cell(5 + i, 2, k).font = Font(bold=True, color=ROAST)
        c = ws.cell(5 + i, 3, v)
        c.alignment = wrap
        ws.row_dimensions[5 + i].height = 32

    # Prospects
    p = wb.create_sheet("Prospects")
    heads = ["ID", "Business", "Type", "Town", "Miles", "Address", "Contact", "Role", "Email", "Email note", "Phone",
             "Website", "Coffee now", "Coffee (0-3)", "Volume (0-2)", "Distance (1-2)", "Local fit (0-2)",
             "Open to switch (0-1)", "Score", "Tier", "Why it fits", "Suggested first line", "Source", "Source URL",
             "Checked", "Status", "Next step / notes"]
    widths = [8, 30, 15, 14, 7, 30, 20, 18, 34, 22, 11, 32, 24, 9, 9, 9, 9, 9, 8, 6, 52, 60, 20, 40, 11, 16, 30]
    p.append(heads)
    for i, r in enumerate(ROWS):
        n = i + 2
        p.append([r["id"], r["name"], r["type"], r["town"], r["miles"], r["address"], r["contact"], r["role"],
                  r["email"] or None, r["email_note"], r["phone"], r["website"], r["coffee_now"], r["c_coffee"],
                  r["c_volume"], "=IF(E%d<=12,2,1)" % n, r["c_local"], r["c_open"],
                  "=SUM(N%d:R%d)" % (n, n), '=IF(S%d>=8,"A",IF(S%d>=6,"B","C"))' % (n, n),
                  r["why"], r["opener"], r["source"], r["source_url"], r["checked"], r["status"], None])
    for i, w in enumerate(widths):
        p.column_dimensions[L(i + 1)].width = w
    for c in p[1]:
        c.font, c.fill = head_font, head_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    p.row_dimensions[1].height = 34
    last = N + 1
    for row in p.iter_rows(min_row=2, max_row=last):
        for c in row:
            c.alignment = wrap
            c.border = Border(bottom=thin)
    for n in range(2, last + 1):
        p.row_dimensions[n].height = 46
        for col in ("L", "X"):
            c = p["%s%d" % (col, n)]
            c.hyperlink = c.value
            c.font = Font(color="1F5F8B", underline="single")
    t = Table(displayName="Prospects", ref="A1:%s%d" % (L(len(heads)), last))
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    p.add_table(t)
    p.freeze_panes = "C2"
    dv = DataValidation(type="list", formula1="=Lists!$A$2:$A$%d" % (len(D.STATUS) + 1), allow_blank=True)
    p.add_data_validation(dv)
    dv.add("Z2:Z%d" % last)
    for tier, color in (("A", "D7EAD9"), ("B", "F6E7C8"), ("C", "EDE7DF")):
        p.conditional_formatting.add("T2:T%d" % last, CellIsRule(operator="equal", formula=['"%s"' % tier],
                                     fill=PatternFill("solid", fgColor=color), font=Font(bold=True)))
    p.conditional_formatting.add("Z2:Z%d" % last, FormulaRule(formula=['$Z2="Won"'],
                                 fill=PatternFill("solid", fgColor="D7EAD9")))

    # Summary
    s = wb.create_sheet("Summary")
    s.column_dimensions["A"].width = 22
    for col in "BCDE":
        s.column_dimensions[col].width = 11
    s.append(["By type", "Tier A", "Tier B", "Tier C", "Total"])
    for t_ in D.TYPES:
        n = s.max_row + 1
        s.append([t_] + ['=COUNTIFS(Prospects!$C$2:$C$%d,$A%d,Prospects!$T$2:$T$%d,"%s")' % (last, n, last, k)
                         for k in "ABC"] + ["=SUM(B%d:D%d)" % (n, n)])
    s.append(["Total"] + ["=SUM(%s2:%s%d)" % (c, c, s.max_row) for c in "BCDE"])
    s.append([])
    top = s.max_row + 1
    s.append(["By town", "Miles", "Prospects", "Tier A"])
    for town, miles in D.TOWNS.items():
        n = s.max_row + 1
        s.append([town, miles, "=COUNTIF(Prospects!$D$2:$D$%d,$A%d)" % (last, n),
                  '=COUNTIFS(Prospects!$D$2:$D$%d,$A%d,Prospects!$T$2:$T$%d,"A")' % (last, n, last)])
    for r_ in (1, top):
        for c in s[r_]:
            c.font, c.fill = head_font, head_fill

    # Scoring
    sc = wb.create_sheet("Scoring")
    sc.column_dimensions["A"].width = 22
    sc.column_dimensions["B"].width = 10
    sc.column_dimensions["C"].width = 90
    sc.append(["Part", "Points", "How it's judged"])
    for part in SCORING:
        sc.append([part[0], part[1], part[2]])
    sc.append([])
    sc.append(["Tier A", "8 to 10", "Call or visit first."])
    sc.append(["Tier B", "6 to 7", "Email first, then call."])
    sc.append(["Tier C", "4 to 5", "Keep for a later round."])
    for c in sc[1]:
        c.font, c.fill = head_font, head_fill
    for row in sc.iter_rows(min_row=2):
        for c in row:
            c.alignment = wrap

    # Skipped, Duplicates
    for title, heads_, rows_, w_ in (
        ("Skipped", ["Place", "Why it's left off", "Rule"], D.SKIPPED, [36, 64, 30]),
        ("Duplicates", ["Kept", "Duplicate seen", "How it was caught"], D.DUPES, [30, 40, 44]),
    ):
        x = wb.create_sheet(title)
        x.append(heads_)
        for r_ in rows_:
            x.append(list(r_))
        for i, w in enumerate(w_):
            x.column_dimensions[L(i + 1)].width = w
        for c in x[1]:
            c.font, c.fill = head_font, head_fill
        for row in x.iter_rows(min_row=2):
            for c in row:
                c.alignment = wrap

    # Lists
    li = wb.create_sheet("Lists")
    li.append(["Status"])
    for st in D.STATUS:
        li.append([st])
    li.column_dimensions["A"].width = 20
    li.sheet_state = "visible"
    wb.active = 1
    wb.save(os.path.join(DOCS, XLSX))


SCORING = [
    ("Coffee", "0 to 3", "3 if espresso or coffee is a core part of the menu, 2 if drip or self-serve coffee is offered, 1 if coffee is a side offering or doesn't exist yet."),
    ("Volume", "0 to 2", "2 for long hours, lots of seats or a big headcount; 1 for a small shop or office."),
    ("Distance", "1 to 2", "2 if 12 miles or less from the roastery, 1 if 13 to 25 miles. Over 25 is left off."),
    ("Local fit", "0 to 2", "2 if the menu, site or shelves already name local suppliers, 1 for some sign of it, 0 for none."),
    ("Open to switch", "0 to 1", "1 if there's no named roaster, a distributor blend, pods or a new owner; 0 if they already buy from a local roaster or a contract."),
]


# ----------------------------------------------------------------- page
def tier_badge(t):
    return '<span class="tier tier--%s">%s</span>' % (t.lower(), t)


def row_html(r):
    contact = '<b>%s</b><span>%s</span>' % (e(r["contact"]), e(r["role"]))
    if r["email"]:
        contact += '<a href="mailto:%s">%s</a>' % (e(r["email"]), e(r["email"]))
    else:
        contact += '<span class="muted">Contact form only</span>'
    contact += '<span class="mono">%s</span>' % r["phone"]
    pct = r["score"] * 10
    return (
        '<tr data-type="%s" data-tier="%s" data-town="%s" data-score="%d" data-miles="%d" data-name="%s" data-text="%s">'
        '<td class="c-id mono">%s</td>'
        '<td class="c-biz"><b>%s</b><span class="type">%s</span><span class="addr">%s</span><span class="now">Coffee now: %s</span></td>'
        '<td class="c-town">%s<span class="mono">%d mi</span></td>'
        '<td class="c-contact">%s</td>'
        '<td class="c-fit">%s<span class="score mono">%d<small>/10</small></span><span class="bar" aria-hidden="true"><i style="width:%d%%"></i></span></td>'
        '<td class="c-why"><p>%s</p><p class="opener"><span>First line</span>%s</p></td>'
        '<td class="c-src"><a href="%s" rel="nofollow">%s</a><span class="mono">Checked %s</span></td>'
        '</tr>' % (
            e(r["type"]), r["tier"], e(r["town"]), r["score"], r["miles"], e(r["name"]),
            e((r["name"] + " " + r["contact"] + " " + r["why"] + " " + r["town"] + " " + r["type"]).lower()),
            r["id"], e(r["name"]), e(r["type"]), e(r["address"]), e(r["coffee_now"]),
            e(r["town"]), r["miles"], contact, tier_badge(r["tier"]), r["score"], pct,
            e(r["why"]), e(r["opener"]), e(r["source_url"]), e(r["source"]), r["checked"][5:].replace("-", "/")))


def bars(groups):
    mx = max(sum(v.values()) for v in groups.values())
    out = []
    for label, v in groups.items():
        tot = sum(v.values())
        segs = "".join('<i class="seg seg--%s" style="width:%.2f%%" title="Tier %s: %d"></i>' % (k.lower(), 100.0 * v[k] / mx, k, v[k])
                       for k in "ABC" if v[k])
        out.append('<div class="bars__row"><span class="bars__label">%s</span><span class="bars__track">%s</span><span class="bars__n mono">%d</span></div>'
                   % (e(label), segs, tot))
    return "".join(out)


def page():
    by_type = OrderedDict((t, Counter(r["tier"] for r in ROWS if r["type"] == t)) for t in D.TYPES)
    by_town = OrderedDict(("%s · %d mi" % (t, m), Counter(r["tier"] for r in ROWS if r["town"] == t)) for t, m in D.TOWNS.items())
    xsize = os.path.getsize(os.path.join(DOCS, XLSX)) // 1024
    csize = max(1, os.path.getsize(os.path.join(DOCS, CSVF)) // 1024)
    ex = ROWS[0]
    type_opts = "".join('<option value="%s">%s</option>' % (e(t), e(t)) for t in D.TYPES)
    town_opts = "".join('<option value="%s">%s</option>' % (e(t), e(t)) for t in D.TOWNS)
    scoring_rows = "".join('<tr><th scope="row">%s</th><td class="mono">%s</td><td>%s</td></tr>' % (a, b, c) for a, b, c in SCORING)
    skipped = "".join('<tr><th scope="row">%s</th><td>%s</td><td><span class="rule">%s</span></td></tr>' % (e(a), e(b), e(c.replace("Rule: ", ""))) for a, b, c in D.SKIPPED)
    dupes = "".join('<tr><th scope="row">%s</th><td>%s</td><td>%s</td></tr>' % (e(a), e(b), e(c)) for a, b, c in D.DUPES)
    rows = "".join(row_html(r) for r in sorted(ROWS, key=lambda r: (-r["score"], r["miles"], r["name"])))
    tierA = [r for r in ROWS if r["tier"] == "A"]
    route = OrderedDict()
    for r in sorted(tierA, key=lambda r: r["miles"]):
        route.setdefault(r["town"], []).append(r["name"])
    route_html = "".join('<li><b>%s</b><span>%s</span></li>' % (e(t), e(", ".join(v))) for t, v in route.items())

    return TEMPLATE.format(
        N=N, A=TIERS["A"], B=TIERS["B"], C=TIERS["C"], EMAILS=EMAILS, FORMS=FORMS,
        SKIPN=len(D.SKIPPED), DUPN=len(D.DUPES), XLSX=XLSX, CSVF=CSVF, XSIZE=xsize, CSIZE=csize,
        TYPE_OPTS=type_opts, TOWN_OPTS=town_opts, ROWS=rows, SCORING=scoring_rows, SKIPPED=skipped, DUPES=dupes,
        BY_TYPE=bars(by_type), BY_TOWN=bars(by_town), ROUTE=route_html, RANGE=D.CHECKED_RANGE,
        EXNAME=e(ex["name"]), EX1=ex["c_coffee"], EX2=ex["c_volume"], EX3=ex["c_distance"], EX4=ex["c_local"],
        EX5=ex["c_open"], EXS=ex["score"], EXT=ex["tier"],
    )


TEMPLATE = open(os.path.join(HERE, "page.html"), encoding="utf-8").read()

if __name__ == "__main__":
    write_csv()
    write_xlsx()
    with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
        f.write(page())
    print("rows %d  tiers %s  emails %d  forms %d" % (N, dict(TIERS), EMAILS, FORMS))
