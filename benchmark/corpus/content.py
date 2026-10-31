"""Deterministic, realistic-looking content for the synthetic corpus.

The documents are the annual report of a fictional public body (a regional water authority):
chapters of prose with figures, money, dates and percentages, varied sentence lengths, lists,
tables with merged header cells, figures with captions, footnotes, reviewer comments and
tracked changes. The decks are the matching board briefing. Everything is drawn from one
``random.Random(seed)`` per file, so the same size always gives the same plan.

Plans are plain JSON consumed by the generators (``gen_docx.js``, ``gen_pptx.js``,
``gen_pptx_pythonpptx.py``). Non-ASCII characters in the text (curly quotes, dashes, the pound
and euro signs) are written as escapes so that this source stays ASCII.
"""

from __future__ import annotations

import random

ORG = "Harbourside Regional Water Authority"
SHORT = "the Authority"
YEAR = 2025
EN_DASH, RSQUO, LDQUO, RDQUO, POUND, EURO, NBSP = "\u2013", "\u2019", "\u201c", "\u201d", "\u00a3", "\u20ac", "\u00a0"

DISTRICTS = ["Eastmoor", "Kellbrook", "Saltmarsh", "Upper Wend", "Rivington", "Hollin Vale", "Northgate",
             "Ashby Cross", "Fenwick", "Torrin Bay", "Westerleigh", "Marlow Fields"]
WORKS = ["Kellbrook treatment works", "Fenwick reservoir", "Saltmarsh pumping station", "Hollin Vale aquifer",
         "Northgate service reservoir", "Torrin Bay desalination pilot"]
PROGRAMMES = ["mains renewal programme", "smart metering pilot", "lead pipe replacement scheme",
              "catchment management plan", "drought resilience scheme", "customer assistance tariff",
              "pressure management project", "digital twin of the network", "sludge-to-energy upgrade",
              "apprenticeship programme"]
SUBJECTS = ["Leakage", "Average daily demand", "Operating expenditure", "Customer satisfaction",
            "The number of supply interruptions", "Energy use per megalitre", "Capital spend",
            "The volume of water abstracted", "Compliance with drinking water standards",
            "The backlog of planned maintenance", "Response time to reported bursts",
            "The proportion of metered households", "Carbon emissions from pumping",
            "Debt outstanding at year end", "The number of written complaints"]
VERBS_UP = ["rose", "increased", "climbed", "grew"]
VERBS_DOWN = ["fell", "decreased", "dropped", "declined"]
CAUSES = ["the dry summer", "higher energy prices", "the cold snap in February",
          "a change in the reporting method", "the completion of the Eastmoor trunk main",
          "lower borrowing costs", "staff shortages in the field teams", "new pressure-reducing valves",
          "the regulator" + RSQUO + "s revised methodology", "population growth in the coastal districts",
          "deferred work from the previous year", "an unusually wet autumn"]
ADJ = ["stable", "ahead of plan", "behind schedule", "within budget", "under review", "on track",
       "broadly in line with forecast", "a priority for the board", "subject to external audit"]
ROLES = ["the Board", "the Audit and Risk Committee", "the Customer Panel", "the executive team",
         "the regulator", "district managers", "the environmental agency", "local councils"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
          "October", "November", "December"]
CHAPTERS = ["Strategic overview", "Water supply and demand", "Network performance", "Customer service",
            "Capital investment", "Financial review", "Environment and sustainability", "People and safety",
            "Governance", "Risk management", "Regulatory compliance", "Digital transformation",
            "Community engagement", "Asset health", "Water quality", "Resilience planning",
            "Procurement and supply chain", "Outlook"]
SECTION_TOPICS = ["Context", "Performance in {y}", "Key risks", "Progress against targets", "Lessons learned",
                  "Regional picture", "Spending", "What customers told us", "Next steps", "Method",
                  "Benchmarking", "Delivery partners", "Funding", "Outcomes"]
AUTHORS = [("Amira Haddad", "AH"), ("Tom Bellweather", "TB"), ("Ines Laurent", "IL"), ("Kwame Asante", "KA")]
COMMENTS = ["Can we cite the source for this figure?", "This sentence is too long; consider splitting it.",
            "Check against the audited accounts before publication.", "Is this the final number?",
            "Suggest moving this to the annex.", "The regulator asked us to phrase this differently last year.",
            "Please add the comparison with {y1}.", "Agreed with finance on 12 March.",
            "Do we have consent to name this district?", "Round to one decimal place for consistency.",
            "Link to the risk register entry here.", "Board members will ask about the cost overrun."]


def money(rng: random.Random) -> str:
    v = rng.choice([rng.uniform(0.2, 9.9), rng.uniform(10, 250)])
    cur = POUND if rng.random() < 0.85 else EURO
    if v < 1:
        return f"{cur}{int(v * 1000)}{NBSP}000"
    return f"{cur}{v:.1f} million"


def pct(rng: random.Random) -> str:
    return f"{rng.uniform(0.3, 24.9):.1f}%"


def count(rng: random.Random) -> str:
    return f"{rng.randint(12, 48000):,}"


def date(rng: random.Random, year: int = YEAR) -> str:
    return f"{rng.randint(1, 28)} {rng.choice(MONTHS)} {year}"


def sentence(rng: random.Random) -> str:
    y = YEAR
    d = rng.choice(DISTRICTS)
    templates = [
        lambda: f"{rng.choice(SUBJECTS)} {rng.choice(VERBS_UP + VERBS_DOWN)} by {pct(rng)} in {y}.",
        lambda: f"{rng.choice(SUBJECTS)} {rng.choice(VERBS_UP)} by {pct(rng)} compared with {y - 1}, mainly because of {rng.choice(CAUSES)}.",
        lambda: f"In the {rng.choice(['first', 'second', 'third', 'fourth'])} quarter, {count(rng)} customers in {d} were affected by planned work, against {count(rng)} a year earlier.",
        lambda: f"The {rng.choice(PROGRAMMES)} is {rng.choice(ADJ)}.",
        lambda: f"We spent {money(rng)} on the {rng.choice(PROGRAMMES)}, of which {pct(rng)} went to contractors based in the region.",
        lambda: f"Progress was {rng.choice(['steady', 'slower than planned', 'uneven', 'encouraging'])}.",
        lambda: f"At the {rng.choice(WORKS)}, output averaged {rng.randint(18, 240)} megalitres a day, with a peak of {rng.randint(240, 410)} on {date(rng)}.",
        lambda: f"{rng.choice(ROLES).capitalize()} reviewed the {rng.choice(PROGRAMMES)} on {date(rng)} and asked for a revised plan by {rng.choice(MONTHS)} {y + 1}.",
        lambda: f"Although {rng.choice(SUBJECTS).lower()} {rng.choice(VERBS_DOWN)} for the {rng.choice(['second', 'third', 'fourth'])} year running, the gap to the {y + 5} target ({pct(rng)}) remains significant, and closing it will depend on {rng.choice(CAUSES)} as much as on our own work.",
        lambda: f"This is a {rng.choice(['significant', 'modest', 'welcome', 'disappointing'])} result.",
        lambda: f"The {d} district accounts for {pct(rng)} of all connections but {pct(rng)} of reported bursts (see Table {rng.randint(1, 12)}).",
        lambda: f"Over the period {y - 4}{EN_DASH}{y}, the {rng.choice(PROGRAMMES)} replaced {rng.randint(40, 900)} km of main at an average cost of {POUND}{rng.randint(180, 950)} per metre.",
        lambda: f"As one district manager put it, {LDQUO}{rng.choice(['the work is never finished', 'every metre of pipe has a history', 'customers notice the outages, not the investment', 'we are fixing the 1970s'])}{RDQUO}.",
        lambda: f"We expect {rng.choice(SUBJECTS).lower()} to be {rng.choice(ADJ)} by March {y + 1}, provided that {rng.choice(CAUSES)} does not recur.",
        lambda: f"{rng.choice(ROLES).capitalize()} approved {money(rng)} of additional funding.",
        lambda: f"Of the {count(rng)} complaints received, {pct(rng)} concerned billing, {pct(rng)} water pressure and the remainder taste, odour or discolouration.",
        lambda: f"The target was {rng.choice(['met', 'missed', 'exceeded', 'met in eleven of twelve months'])}.",
        lambda: f"Following the incident at the {rng.choice(WORKS)} on {date(rng)}, which interrupted supply to {count(rng)} properties for up to {rng.randint(3, 40)} hours, we commissioned an independent review whose {rng.randint(6, 31)} recommendations we have accepted in full.",
        lambda: f"Unit costs in {d} were {pct(rng)} {rng.choice(['above', 'below'])} the regional average.",
        lambda: f"Since {MONTHS[rng.randrange(12)]} {y - 2}, {count(rng)} households have joined the {rng.choice(PROGRAMMES)}.",
        lambda: f"Why does this matter?",
        lambda: f"Each {POUND}1 invested in {rng.choice(['leak detection', 'pressure management', 'catchment work', 'metering'])} saved an estimated {POUND}{rng.uniform(1.1, 4.8):.2f} over ten years.",
        lambda: f"The {rng.choice(PROGRAMMES)} and the {rng.choice(PROGRAMMES)} share a project board, a risk register and, from April {y + 1}, a single budget line of {money(rng)}.",
        lambda: f"We did not meet our commitment on {rng.choice(SUBJECTS).lower()}, and we explain why in chapter {rng.randint(2, 9)}.",
    ]
    return rng.choice(templates)()


def paragraph_text(rng: random.Random, lo: int = 2, hi: int = 7) -> str:
    return " ".join(sentence(rng) for _ in range(rng.randint(lo, hi)))


def split_runs(rng: random.Random, text: str) -> list[dict]:
    """Split text into several runs, as Word's edit history leaves it: most runs plain,
    some bold or italic, and adjacent runs with the same formatting (rsid-like splits)."""
    words = text.split(" ")
    n = len(words)
    cuts = sorted(set(rng.sample(range(1, n), k=min(n - 1, rng.randint(1, 6))))) if n > 1 else []
    pieces, start = [], 0
    for c in cuts + [n]:
        pieces.append(" ".join(words[start:c]) + (" " if c < n else ""))
        start = c
    runs = []
    for piece in pieces:
        r = rng.random()
        runs.append({"k": "r", "text": piece, "b": r < 0.08 or 0.97 < r, "i": 0.08 <= r < 0.15 or 0.97 < r})
    # Now and then split a word inside a run without changing the formatting.
    if runs and rng.random() < 0.3:
        j = rng.randrange(len(runs))
        t = runs[j]["text"]
        if len(t) > 6:
            m = rng.randint(2, len(t) - 2)
            runs[j:j + 1] = [dict(runs[j], text=t[:m]), dict(runs[j], text=t[m:])]
    return runs


def table_spec(rng: random.Random, kind: str, n: int) -> dict:
    """kind: 'finance' (two header rows, merged cells), 'simple', or 'wide' (landscape annex)."""
    if kind == "finance":
        rows = [
            {"header": True, "cells": [{"text": "Indicator", "rowSpan": 2}, {"text": str(YEAR - 1), "colSpan": 2},
                                       {"text": str(YEAR), "colSpan": 2}, {"text": "Change", "rowSpan": 2}]},
            {"header": True, "cells": [{"text": "Budget"}, {"text": "Actual"}, {"text": "Budget"}, {"text": "Actual"}]},
        ]
        items = SUBJECTS + [p.capitalize() for p in PROGRAMMES]
        rng.shuffle(items)
        for k in range(n):
            vals = [rng.uniform(1, 400) for _ in range(4)]
            rows.append({"cells": [{"text": items[k % len(items)]}] + [{"text": f"{v:,.1f}", "align": "right"} for v in vals]
                         + [{"text": f"{(vals[3] - vals[1]) / vals[1] * 100:+.1f}%", "align": "right"}]})
        rows.append({"total": True, "cells": [{"text": "Total"}] + [{"text": f"{rng.uniform(800, 3000):,.1f}", "align": "right"} for _ in range(4)] + [{"text": ""}]})
        return {"widths": [3400, 1250, 1250, 1250, 1250, 1100], "rows": rows, "caption": "Spending by indicator, " + POUND + " thousand"}
    if kind == "simple":
        heads = rng.choice([["District", "Connections", "Bursts", "Interruptions (h)", "Complaints"],
                            ["Site", "Output (Ml/d)", "Energy (kWh/Ml)", "Compliance"],
                            ["Programme", "Start", "Finish", "Budget", "Status", "Owner"]])
        rows = [{"header": True, "cells": [{"text": h} for h in heads]}]
        for k in range(n):
            cells = [{"text": rng.choice(DISTRICTS + WORKS)}]
            for h in heads[1:]:
                if h in ("Start", "Finish"):
                    cells.append({"text": f"{rng.choice(MONTHS)[:3]} {rng.randint(YEAR - 3, YEAR + 3)}"})
                elif h == "Status":
                    cells.append({"text": rng.choice(["On track", "At risk", "Complete", "Delayed"])})
                elif h == "Owner":
                    cells.append({"text": rng.choice(["Operations", "Finance", "Capital delivery", "Customer"])})
                elif h == "Compliance":
                    cells.append({"text": f"{rng.uniform(98.5, 100):.2f}%", "align": "right"})
                else:
                    cells.append({"text": f"{rng.randint(3, 9000):,}", "align": "right"})
            rows.append({"cells": cells, "shade": k % 2 == 1})
        total = 9000
        w = [total // len(heads)] * len(heads)
        w[0] += total - sum(w)
        return {"widths": w, "rows": rows, "caption": "Summary by " + heads[0].lower()}
    # wide
    heads = ["District", "Pop. served", "Connections", "Length of main (km)", "Bursts", "Leakage (Ml/d)",
             "Interruptions (h)", "Complaints", "Capital spend (" + POUND + "k)", "Satisfaction"]
    rows = [{"header": True, "cells": [{"text": h} for h in heads]}]
    for k in range(n):
        rows.append({"cells": [{"text": DISTRICTS[k % len(DISTRICTS)] + ("" if k < len(DISTRICTS) else f" {k // len(DISTRICTS) + 1}")}]
                     + [{"text": f"{rng.randint(10, 99999):,}", "align": "right"} for _ in range(len(heads) - 2)]
                     + [{"text": f"{rng.uniform(6.0, 9.6):.1f}", "align": "right"}], "shade": k % 2 == 1})
    return {"widths": [1700] + [1380] * 9, "rows": rows, "caption": "District statistics"}


SIZES_DOCX = {
    # chapters, sections per chapter (range); calibrated so that LibreOffice lays out ~5/50/200 A4 pages.
    "S": {"chapters": 1, "sections": (2, 2), "pages": 5},
    "M": {"chapters": 10, "sections": (3, 4), "pages": 50},
    "L": {"chapters": 37, "sections": (3, 5), "pages": 200},
}


def docx_plan(size: str, images: list[dict], clean: bool = False, seed: int | None = None) -> dict:
    cfg = SIZES_DOCX[size]
    rng = random.Random(seed if seed is not None else {"S": 11, "M": 22, "L": 33}[size])
    target_pages = cfg["pages"]
    body: list[dict] = []
    footnotes: dict[str, str] = {}
    fig = tab = 0
    img_i = 0
    list_instance = 0

    def add_para():
        body.append({"t": "p", "inl": split_runs(rng, paragraph_text(rng))})

    body.append({"t": "title", "text": f"{ORG}"})
    body.append({"t": "subtitle", "text": f"Annual Report and Accounts {YEAR}{EN_DASH}{str(YEAR + 1)[2:]}"})
    body.append({"t": "p", "inl": [{"k": "r", "text": f"Published {date(rng, YEAR + 1)}. Presented to the Board under section 12 of the Water Services Act (fictional).", "i": True}]})
    body.append({"t": "pagebreak"})
    body.append({"t": "toc"})
    body.append({"t": "pagebreak"})

    for ch in range(cfg["chapters"]):
        name = CHAPTERS[ch % len(CHAPTERS)]
        if ch >= len(CHAPTERS):
            name += f" {EN_DASH} {DISTRICTS[ch % len(DISTRICTS)]}"
        body.append({"t": "h", "level": 1, "text": f"{ch + 1}. {name}"})
        add_para()
        for s in range(rng.randint(*cfg["sections"])):
            topic = rng.choice(SECTION_TOPICS).format(y=YEAR)
            body.append({"t": "h", "level": 2, "text": f"{ch + 1}.{s + 1} {topic}"})
            for _ in range(rng.randint(2, 4)):
                add_para()
            r = rng.random()
            if r < 0.35:
                list_instance += 1
                kind = "bullet" if rng.random() < 0.65 else "num"
                for _ in range(rng.randint(3, 6)):
                    body.append({"t": "li", "list": kind, "level": 0, "instance": list_instance,
                                 "inl": split_runs(rng, sentence(rng))})
                    if kind == "bullet" and rng.random() < 0.2:
                        body.append({"t": "li", "list": kind, "level": 1, "instance": list_instance,
                                     "inl": [{"k": "r", "text": sentence(rng)}]})
            if rng.random() < 0.55:
                for h3 in range(rng.randint(1, 2)):
                    body.append({"t": "h", "level": 3, "text": f"{rng.choice(DISTRICTS)} {rng.choice(['district', 'zone', 'catchment'])}"})
                    for _ in range(rng.randint(1, 3)):
                        add_para()
            if rng.random() < 0.40 or (ch == 0 and s == 0 and tab == 0):  # every file has a table
                tab += 1
                kind = "finance" if rng.random() < 0.5 else "simple"
                n = rng.randint(10, 18) if kind == "finance" else rng.randint(5, 12)
                spec = table_spec(rng, kind, n)
                body.append({"t": "caption", "text": f"Table {tab}: {spec['caption']}"})
                body.append({"t": "table", **spec})
                add_para()
            if (rng.random() < 0.30 or (ch == 0 and s == 1 and fig == 0)) and images:  # and a figure
                fig += 1
                im = images[img_i % len(images)]
                img_i += 1
                w_px = 600 if im["kind"] != "photo" else 420
                body.append({"t": "img", "path": im["path"], "w": w_px, "h": round(w_px * im["h"] / im["w"]),
                             "alt": f"Figure {fig}"})
                body.append({"t": "caption", "text": f"Figure {fig}: {rng.choice(['Monthly demand against forecast', 'Capital spend by district', 'Supply mix by source', 'Bursts by zone and month', 'Pressure against age of main', 'Site survey'])}, {YEAR - 4}{EN_DASH}{YEAR}"})
            add_para()

    # Landscape annex: the second section.
    annex = [{"t": "h", "level": 1, "text": "Annex A. District statistics"}]
    annex.append({"t": "p", "inl": split_runs(rng, paragraph_text(rng, 2, 3))})
    for k in range(max(1, target_pages // 25)):
        tab += 1
        spec = table_spec(rng, "wide", rng.randint(10, 14) if size == "S" else rng.randint(14, 24))
        annex.append({"t": "caption", "text": f"Table {tab}: {spec['caption']} (part {k + 1})"})
        annex.append({"t": "table", **spec})
        annex.append({"t": "p", "inl": split_runs(rng, paragraph_text(rng, 1, 3))})

    # Footnotes (~1 per 3 pages), comments (~1 per 2 pages), tracked changes (~1 per page),
    # spread evenly over the body paragraphs of the main section.
    paras = [b for b in body if b["t"] == "p" and len(b["inl"]) >= 2][1:]
    def spread(n, offset):
        if not paras or n <= 0:
            return []
        step = len(paras) / n
        return [paras[int(offset * step + k * step) % len(paras)] for k in range(n)]

    for k, p in enumerate(spread(max(1, target_pages // 3), 0.3)):
        fid = str(k + 1)
        footnotes[fid] = rng.choice([
            f"Figures for {YEAR - 1} restated following the audit of {date(rng)}.",
            "Source: internal asset register, extracted at year end.",
            f"Excludes the {rng.choice(WORKS)}, which was transferred to the Authority during the year.",
            "Regulatory definition; see the glossary in Annex B.",
            f"Prices in {YEAR} terms, deflated with the consumer price index.",
        ])
        p["inl"].insert(rng.randint(1, len(p["inl"])), {"k": "fn", "id": int(fid)})

    comments = []
    rev_id = 5000
    if not clean:
        for k, p in enumerate(spread(max(1, target_pages // 2), 0.55)):
            a, ini = AUTHORS[k % len(AUTHORS)]
            comments.append({"id": k, "author": a, "initials": ini,
                             "date": f"{YEAR + 1}-0{1 + k % 3}-{1 + k % 27:02d}T{9 + k % 8:02d}:{(k * 7) % 60:02d}:00Z",
                             "text": rng.choice(COMMENTS).format(y1=YEAR - 1)})
            # Anchor over one or two runs of the paragraph.
            runs_idx = [i for i, r in enumerate(p["inl"]) if r["k"] == "r"]
            i0 = rng.choice(runs_idx[:-1] or runs_idx)
            i1 = min(i0 + rng.randint(0, 1), len(p["inl"]) - 1)
            p["inl"].insert(i1 + 1, {"k": "ce", "id": k})
            p["inl"].insert(i0, {"k": "cs", "id": k})
        for k, p in enumerate(spread(target_pages, 0.05)):
            a, _ = AUTHORS[(k + 1) % len(AUTHORS)]
            when = f"{YEAR + 1}-02-{1 + k % 27:02d}T{8 + k % 9:02d}:{(k * 13) % 60:02d}:00Z"
            mode = k % 3  # 0 insertion, 1 deletion, 2 replacement (deletion then insertion)
            runs_idx = [i for i, r in enumerate(p["inl"]) if r["k"] == "r" and len(r["text"].split()) >= 3]
            if not runs_idx:
                continue
            i = rng.choice(runs_idx)
            r = p["inl"][i]
            words = r["text"].split(" ")
            j = rng.randint(1, max(1, len(words) - 2))
            before, mid, after = " ".join(words[:j]) + " ", words[j] + " ", " ".join(words[j + 1:])
            fmt = {"b": r.get("b", False), "i": r.get("i", False)}
            new = [{"k": "r", "text": before, **fmt}]
            if mode in (1, 2):
                rev_id += 1
                new.append({"k": "del", "text": mid, "id": rev_id, "author": a, "date": when, **fmt})
            if mode in (0, 2):
                rev_id += 1
                ins_text = rng.choice(["approximately ", "in total ", "provisional ", "audited ", "at least ",
                                       "for the full year ", "(unaudited) ", "again "])
                if mode == 2:
                    ins_text = rng.choice(["revised ", "updated ", "final ", "corrected "])
                new.append({"k": "ins", "text": ins_text, "id": rev_id, "author": a, "date": when, **fmt})
            if mode == 0:
                new.append({"k": "r", "text": mid, **fmt})
            if after:
                new.append({"k": "r", "text": after, **fmt})
            p["inl"][i:i + 1] = new

    return {
        "meta": {"title": f"{ORG} annual report {YEAR}", "creator": "Corporate Affairs", "size": size,
                 "header": f"{ORG} {EN_DASH} Annual Report {YEAR}", "clean": clean},
        "sections": [{"orientation": "portrait", "blocks": body}, {"orientation": "landscape", "blocks": annex}],
        "footnotes": footnotes,
        "comments": comments,
    }


# ------------------------------------------------------------------------------------------
# Decks
SIZES_PPTX = {"S": 15, "M": 60, "L": 150}


def deck_plan(size: str, images: list[dict], seed: int | None = None) -> dict:
    n = SIZES_PPTX[size]
    rng = random.Random(seed if seed is not None else {"S": 101, "M": 202, "L": 303}[size])
    slides: list[dict] = []
    n_sections = max(2, n // 12)
    sec_names = [CHAPTERS[k % len(CHAPTERS)] + ("" if k < len(CHAPTERS) else f" ({DISTRICTS[k % len(DISTRICTS)]})")
                 for k in range(n_sections)]

    def notes():
        return paragraph_text(rng, 2, 5)

    slides.append({"type": "title", "section": "Introduction", "title": f"{ORG}",
                   "subtitle": f"Board briefing {EN_DASH} performance in {YEAR}", "notes": notes()})
    slides.append({"type": "bullets", "section": "Introduction", "title": "Agenda",
                   "bullets": [s for s in sec_names[:8]], "notes": notes()})
    # Content slides cycle through kinds; section dividers every n/n_sections slides.
    kinds = ["bullets_pic", "chart", "table", "two_col", "chart", "stats", "long_text", "chart", "bullets_pic", "table"]
    chart_types = ["bar", "line", "pie"]
    per_sec = (n - 3) / n_sections
    k_chart = k_img = 0
    i = 0
    for s in range(n_sections):
        sec = sec_names[s]
        slides.append({"type": "section", "section": sec, "title": sec, "subtitle": f"Section {s + 1}", "notes": notes()})
        count_here = int(round(per_sec * (s + 1))) - int(round(per_sec * s)) - 1
        for _ in range(count_here):
            kind = kinds[i % len(kinds)]
            i += 1
            title = rng.choice([f"{rng.choice(SUBJECTS)} in {rng.choice(DISTRICTS)}", f"The {rng.choice(PROGRAMMES)}",
                                f"{sec}: {rng.choice(SECTION_TOPICS).format(y=YEAR).lower()}",
                                f"What changed since {YEAR - 1}"])
            sl = {"type": kind, "section": sec, "title": title, "notes": notes()}
            if kind in ("bullets_pic", "two_col"):
                sl["bullets"] = [sentence(rng) for _ in range(rng.randint(3, 5))]
                if kind == "two_col":
                    sl["right"] = [sentence(rng) for _ in range(rng.randint(3, 5))]
                    sl["left_head"], sl["right_head"] = rng.choice([("Achieved", "Outstanding"), (str(YEAR - 1), str(YEAR)),
                                                                     ("Risks", "Mitigations")])
                else:
                    im = images[k_img % len(images)]
                    k_img += 1
                    sl["image"] = {"path": im["path"], "w": im["w"], "h": im["h"]}
            elif kind == "chart":
                ct = chart_types[k_chart % 3]
                k_chart += 1
                if ct == "pie":
                    labels = rng.sample(["Billing", "Pressure", "Taste/odour", "Discolouration", "Metering", "Other"], 5)
                    sl["chart"] = {"type": "pie", "series": [{"name": "Share of complaints", "labels": labels,
                                                              "values": [rng.randint(5, 40) for _ in labels]}]}
                elif ct == "bar":
                    labels = rng.sample(DISTRICTS, rng.randint(5, 8))
                    sl["chart"] = {"type": "bar", "series": [
                        {"name": str(YEAR - 1), "labels": labels, "values": [round(rng.uniform(2, 30), 1) for _ in labels]},
                        {"name": str(YEAR), "labels": labels, "values": [round(rng.uniform(2, 30), 1) for _ in labels]}]}
                else:
                    labels = [m[:3] for m in MONTHS]
                    sl["chart"] = {"type": "line", "series": [
                        {"name": "Demand", "labels": labels, "values": [round(rng.uniform(180, 260), 1) for _ in labels]},
                        {"name": "Forecast", "labels": labels, "values": [round(rng.uniform(190, 250), 1) for _ in labels]}]}
                sl["chart"]["title"] = rng.choice(["Bursts per 1,000 km", "Monthly demand (Ml/d)", "Complaints by type",
                                                   "Capital spend (" + POUND + "m)"])
                sl["commentary"] = [sentence(rng) for _ in range(rng.randint(2, 3))]
            elif kind == "table":
                cols = rng.randint(4, 6)
                rows = rng.randint(5, 8)
                heads = ["District", "Connections", "Bursts", "Leakage (Ml/d)", "Complaints", "Spend (" + POUND + "k)"][:cols]
                body = [[rng.choice(DISTRICTS)] + [f"{rng.randint(3, 9000):,}" for _ in range(cols - 1)] for _ in range(rows - 1)]
                sl["table"] = [heads] + body
            elif kind == "stats":
                sl["stats"] = [{"value": pct(rng), "label": rng.choice(SUBJECTS)} for _ in range(3)]
                sl["caption"] = sentence(rng)
            elif kind == "long_text":
                sl["text"] = [paragraph_text(rng, 4, 6) for _ in range(3)]
            slides.append(sl)
    slides.append({"type": "closing", "section": sec_names[-1], "title": "Next steps",
                   "bullets": [sentence(rng) for _ in range(4)], "notes": notes()})
    assert len(slides) == n, (len(slides), n)
    return {"meta": {"title": f"{ORG} board briefing {YEAR}", "size": size, "company": ORG,
                     "footer": f"{ORG} {EN_DASH} Board briefing {YEAR}"}, "slides": slides}
