#!/usr/bin/env python3
"""
job_tracker.py — pulls product roles straight from ATS APIs and maintains
a self-updating spreadsheet.

Skips Google entirely: Greenhouse, Lever, Ashby and Workday each expose the
JSON their own job boards are built on, so postings are live rather than
whatever the crawler last indexed.

  python3 job_tracker.py run                 fetch + refresh the workbook
  python3 job_tracker.py add-board <url>     register a board from its URL
  python3 job_tracker.py boards              list configured boards
"""
import argparse, json, os, re, sys, datetime as dt
import urllib.request, urllib.error
from concurrent.futures import ThreadPoolExecutor

HERE   = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, "config.json")
STATE  = os.path.join(HERE, "state.json")
XLSX   = os.path.join(HERE, "Job Tracker.xlsx")

TODAY  = dt.date.today()
UA     = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 job-tracker"

# Columns the human owns. Never overwritten by a fetch.
TRACK   = ["Status", "Priority", "Applied On", "Notes"]
STATUS  = ["Interested", "Applied", "Screen", "Interview", "Onsite",
           "Offer", "Rejected", "Passed", "Closed"]
PRIORITY = ["A", "B", "C"]


# ---------------------------------------------------------------- http

def _http(url, payload=None, timeout=25, tries=3):
    data = None
    hdrs = {"User-Agent": UA, "Accept": "application/json"}
    if payload is not None:
        data = json.dumps(payload).encode()
        hdrs["Content-Type"] = "application/json"
    last = None
    for n in range(tries):
        try:
            req = urllib.request.Request(url, data=data, headers=hdrs)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8", "replace"))
        except Exception as e:                       # noqa: BLE001
            last = e
    raise last


def _iso(s):
    """Best-effort ISO-8601 -> date."""
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(str(s).replace("Z", "+00:00")).date()
    except Exception:                                # noqa: BLE001
        return None


def _workday_date(text):
    """'Posted 3 Days Ago' / 'Posted Today' / 'Posted 30+ Days Ago' -> date."""
    if not text:
        return None
    t = text.lower()
    if "today" in t:
        return TODAY
    if "yesterday" in t:
        return TODAY - dt.timedelta(days=1)
    m = re.search(r"(\d+)\+?\s*day", t)
    if m:
        return TODAY - dt.timedelta(days=int(m.group(1)))
    m = re.search(r"(\d+)\+?\s*month", t)
    if m:
        return TODAY - dt.timedelta(days=30 * int(m.group(1)))
    return None


# ---------------------------------------------------------------- fetchers
# Each returns (postings, error_or_None).

def fetch_greenhouse(tok, name):
    url = f"https://boards-api.greenhouse.io/v1/boards/{tok}/jobs?content=false"
    d = _http(url)
    out = []
    for j in d.get("jobs", []):
        posted = _iso(j.get("first_published")) or _iso(j.get("updated_at"))
        out.append({
            "uid":      f"greenhouse:{tok}:{j.get('id')}",
            "source":   "Greenhouse",
            "company":  name or j.get("company_name") or tok,
            "title":    (j.get("title") or "").strip(),
            "location": ((j.get("location") or {}).get("name") or "").strip(),
            "url":      j.get("absolute_url") or "",
            "posted":   posted,
            "dept":     ", ".join(x.get("name", "") for x in (j.get("departments") or [])),
            "req":      j.get("requisition_id") or "",
        })
    return out


def fetch_lever(tok, name):
    d = _http(f"https://api.lever.co/v0/postings/{tok}?mode=json")
    out = []
    for j in d if isinstance(d, list) else []:
        cats = j.get("categories") or {}
        locs = cats.get("allLocations") or ([cats["location"]] if cats.get("location") else [])
        created = j.get("createdAt")
        posted = None
        if created:
            try:
                posted = dt.datetime.utcfromtimestamp(int(created) / 1000).date()
            except Exception:                        # noqa: BLE001
                posted = None
        out.append({
            "uid":      f"lever:{tok}:{j.get('id')}",
            "source":   "Lever",
            "company":  name or tok,
            "title":    (j.get("text") or "").strip(),
            "location": " / ".join(locs),
            "url":      j.get("hostedUrl") or "",
            "posted":   posted,
            "dept":     cats.get("team") or "",
            "req":      "",
        })
    return out


def fetch_ashby(tok, name):
    d = _http(f"https://api.ashbyhq.com/posting-api/job-board/{tok}")
    out = []
    for j in d.get("jobs", []):
        if j.get("isListed") is False:
            continue
        locs = [j.get("location") or ""] + [
            s.get("location", "") for s in (j.get("secondaryLocations") or [])
        ]
        locs = [x for x in locs if x]
        if j.get("isRemote") and not any("remote" in x.lower() for x in locs):
            locs.append("Remote")
        out.append({
            "uid":      f"ashby:{tok}:{j.get('id')}",
            "source":   "Ashby",
            "company":  name or tok,
            "title":    (j.get("title") or "").strip(),
            "location": " / ".join(locs),
            "url":      j.get("jobUrl") or j.get("applyUrl") or "",
            "posted":   _iso(j.get("publishedAt")),
            "dept":     j.get("department") or j.get("team") or "",
            "req":      "",
        })
    return out


def _wd_page(api, offset, facets=None):
    return _http(api, {"appliedFacets": facets or {}, "limit": 20,
                       "offset": offset, "searchText": ""})


def _wd_norm(j, tok, name, base, site):
    path = j.get("externalPath") or ""
    return {
        "uid":      f"workday:{tok}:{path}",
        "source":   "Workday",
        "company":  name,
        "title":    (j.get("title") or "").strip(),
        "location": (j.get("locationsText") or "").strip(),
        "url":      f"{base}/{site}{path}",
        "posted":   _workday_date(j.get("postedOn")),
        "dept":     j.get("timeType") or "",
        "req":      ", ".join(j.get("bulletFields") or []),
    }


def fetch_workday(tok, meta, max_age_days=30, page_cap=25, cat_keywords=("product",)):
    """Two passes, unioned.

    1. Category pass — Workday exposes a `jobFamilyGroup` facet per tenant.
       Filtering to the product-ish categories turns CVS's 19,000 postings
       into ~55, and it reaches roles far older than any date window.
    2. Recency pass — results are sorted newest-first, so page until a whole
       page falls outside the window. Covers tenants with no product
       category, and anything filed under the wrong one.

    Note: Workday only populates `total` on the first response; later pages
    report 0, so paging must never be driven off it.
    """
    dc, site = meta.get("dc", 1), meta["site"]
    name = meta.get("name", tok)
    base = f"https://{tok}.wd{dc}.myworkdayjobs.com"
    api  = f"{base}/wday/cxs/{tok}/{site}/jobs"
    cutoff = TODAY - dt.timedelta(days=max_age_days)
    found = {}

    first = _wd_page(api, 0)
    total = first.get("total") or 0
    for j in first.get("jobPostings") or []:
        p = _wd_norm(j, tok, name, base, site)
        found[p["uid"]] = p

    # ---- 1. category pass
    cat_ids = []
    for f in first.get("facets") or []:
        if f.get("facetParameter") != "jobFamilyGroup":
            continue
        for v in f.get("values") or []:
            d = (v.get("descriptor") or "").lower()
            if any(k in d for k in cat_keywords):
                cat_ids.append(v.get("id"))
    for cid in cat_ids[:4]:
        off, cat_total = 0, None
        for _ in range(25):
            d = _wd_page(api, off, {"jobFamilyGroup": [cid]})
            if cat_total is None:
                cat_total = d.get("total") or 0
            posts = d.get("jobPostings") or []
            if not posts:
                break
            for j in posts:
                p = _wd_norm(j, tok, name, base, site)
                found.setdefault(p["uid"], p)
            off += 20
            if off >= cat_total:
                break

    # ---- 2. recency pass
    off = 20
    for _ in range(page_cap):
        if total and off >= total:
            break
        d = _wd_page(api, off)
        posts = d.get("jobPostings") or []
        if not posts:
            break
        dated = [_workday_date(j.get("postedOn")) for j in posts]
        for j, when in zip(posts, dated):
            if when and when < cutoff:
                continue
            p = _wd_norm(j, tok, name, base, site)
            found.setdefault(p["uid"], p)
        # stop only when the entire page is out of window
        if dated and all(w is not None and w < cutoff for w in dated):
            break
        off += 20

    return list(found.values())


# ---------------------------------------------------------------- matching

def build_matcher(m):
    inc = [re.compile(re.escape(t), re.I) for t in m.get("title_include", [])]
    exc = [re.compile(r"\b" + re.escape(t) + r"\b", re.I) for t in m.get("title_exclude", [])]
    locs = [l.lower() for l in m.get("locations_any", [])]
    kws  = [k.lower() for k in m.get("keywords_any", [])]
    max_age = int(m.get("max_age_days", 45))

    def match(p):
        title = p["title"]
        if not title:
            return False, "no title"
        if inc and not any(r.search(title) for r in inc):
            return False, "title"
        if any(r.search(title) for r in exc):
            return False, "excluded"
        if locs:
            loc = (p["location"] or "").lower().strip()
            # Workday collapses multi-site reqs to "5 Locations" and never says
            # which ones. Treat that as unknown and let it through rather than
            # dropping reqs that may well include the ones being looked for.
            unknown = (not loc) or re.fullmatch(r"\d+\s+locations?", loc)
            if not unknown and not any(l in loc for l in locs):
                return False, "location"
        if kws:
            blob = f"{title} {p.get('dept','')}".lower()
            if not any(k in blob for k in kws):
                return False, "keyword"
        if p["posted"] and (TODAY - p["posted"]).days > max_age:
            return False, "stale"
        return True, ""
    return match


# ---------------------------------------------------------------- state

def load_state():
    if os.path.exists(STATE):
        with open(STATE) as f:
            return json.load(f)
    return {"postings": {}, "last_run": None}


def save_state(s):
    with open(STATE, "w") as f:
        json.dump(s, f, indent=1, sort_keys=True)


def read_user_edits():
    """Pull the human-owned columns back out of the workbook so a refresh
    never clobbers what was typed into it."""
    if not os.path.exists(XLSX):
        return {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(XLSX, read_only=True, data_only=True)
    except Exception:                                # noqa: BLE001
        return {}
    edits = {}
    for sheet in ("Active", "Closed"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        try:
            hdr = list(next(rows))
        except StopIteration:
            continue
        if "UID" not in hdr:
            continue
        iu = hdr.index("UID")
        idx = {c: hdr.index(c) for c in TRACK if c in hdr}
        for r in rows:
            if iu >= len(r) or not r[iu]:
                continue
            # Record every tracked column, blanks included: a cleared cell is
            # a deliberate edit, and skipping it would resurrect the old value
            # from state on the next run.
            rec = {}
            for c, i in idx.items():
                v = r[i] if i < len(r) else None
                if isinstance(v, (dt.date, dt.datetime)):
                    v = v.strftime("%Y-%m-%d")
                rec[c] = "" if v is None else str(v).strip()
            edits[str(r[iu])] = rec
    wb.close()
    return edits


# ---------------------------------------------------------------- workbook

def write_workbook(active, closed, boards, cfg, stats):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule

    NAVY  = "1F3550"
    CREAM = "F7F4EE"
    RULE  = "D9D2C5"
    NEW   = "FFF3C4"

    wb = Workbook()
    head_font = Font(bold=True, color="FFFFFF", size=11)
    head_fill = PatternFill("solid", fgColor=NAVY)
    thin = Side(style="thin", color=RULE)
    border = Border(bottom=thin)

    def style_sheet(ws, headers, widths, rows, link_col=None, freeze="A2"):
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(vertical="center", horizontal="left")
        ws.row_dimensions[1].height = 24
        for r in rows:
            ws.append(r)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=False)
        if link_col and ws.max_row > 1:
            li = headers.index(link_col) + 1
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=li)
                if cell.value:
                    cell.hyperlink = cell.value
                    cell.value = "Open posting"
                    cell.font = Font(color="1155CC", underline="single")
        if ws.max_row > 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
        ws.freeze_panes = freeze
        return ws

    # ---- Active -----------------------------------------------------
    hdr = TRACK + ["Title", "Company", "Location", "Posted", "Age (d)",
                   "First Seen", "Source", "Dept", "Req", "Link", "UID"]
    rows = []
    for p in active:
        rows.append([
            p.get("Status", ""), p.get("Priority", ""),
            p.get("Applied On", ""), p.get("Notes", ""),
            p["title"], p["company"], p["location"],
            p["posted"].strftime("%Y-%m-%d") if p["posted"] else "",
            (TODAY - p["posted"]).days if p["posted"] else "",
            p["first_seen"], p["source"], p["dept"], p["req"],
            p["url"], p["uid"],
        ])
    ws = style_sheet(
        wb.active, hdr,
        [13, 9, 12, 34, 46, 22, 30, 11, 8, 11, 12, 20, 14, 14, 40],
        rows, link_col="Link")
    ws.title = "Active"

    n = ws.max_row
    if n > 1:
        dv = DataValidation(type="list", formula1='"%s"' % ",".join(STATUS), allow_blank=True)
        ws.add_data_validation(dv); dv.add(f"A2:A{n}")
        dvp = DataValidation(type="list", formula1='"%s"' % ",".join(PRIORITY), allow_blank=True)
        ws.add_data_validation(dvp); dvp.add(f"B2:B{n}")
        rng = f"A2:O{n}"
        # brand-new since last run
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'$J2="{TODAY.isoformat()}"'], fill=PatternFill("solid", fgColor=NEW)))
        # already applied -> fade
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=['$A2="Applied"'], fill=PatternFill("solid", fgColor=CREAM)))
    ws.column_dimensions["O"].hidden = True

    # ---- New This Run ------------------------------------------------
    fresh = [p for p in active if p["first_seen"] == TODAY.isoformat()]
    style_sheet(wb.create_sheet("New This Run"),
                ["Title", "Company", "Location", "Posted", "Source", "Link"],
                [46, 24, 32, 12, 12, 14],
                [[p["title"], p["company"], p["location"],
                  p["posted"].strftime("%Y-%m-%d") if p["posted"] else "",
                  p["source"], p["url"]] for p in fresh],
                link_col="Link")

    # ---- Closed ------------------------------------------------------
    style_sheet(wb.create_sheet("Closed"),
                TRACK + ["Title", "Company", "Location", "Reason", "First Seen",
                         "Last Seen", "Days Live", "Source", "Link", "UID"],
                [13, 9, 12, 30, 46, 24, 30, 22, 11, 11, 10, 12, 14, 40],
                [[c.get("Status", ""), c.get("Priority", ""), c.get("Applied On", ""),
                  c.get("Notes", ""), c["title"], c["company"], c["location"],
                  c.get("reason", ""), c["first_seen"], c["last_seen"],
                  c.get("days_live", ""), c["source"], c["url"], c["uid"]]
                 for c in closed],
                link_col="Link")

    # ---- Boards ------------------------------------------------------
    style_sheet(wb.create_sheet("Boards"),
                ["Company", "ATS", "Token", "Total Live", "Matched", "Status"],
                [30, 14, 26, 12, 10, 44], boards)

    # ---- Search Queries ---------------------------------------------
    since = (TODAY - dt.timedelta(days=7)).isoformat()
    qrows = []
    for q in cfg.get("google_queries", []):
        full = f"{q} after:{since}"
        qrows.append([q, full,
                      "https://www.google.com/search?q=" + urllib.request.quote(full)])
    wsq = style_sheet(wb.create_sheet("Search Queries"),
                      ["Query", "With freshness filter", "Run"],
                      [64, 74, 14], qrows, link_col="Run")
    wsq.cell(row=wsq.max_row + 2, column=1,
             value="Manual backstop for boards not in the Boards tab. "
                   "The Active tab is pulled live from each ATS API, so it does not "
                   "depend on Google's index.").font = Font(italic=True, color="777777")

    # ---- Run Log -----------------------------------------------------
    style_sheet(wb.create_sheet("Run Log"), ["Metric", "Value"], [34, 60],
                [[k, v] for k, v in stats.items()])

    wb.save(XLSX)


# ---------------------------------------------------------------- run

def cmd_run(args):
    cfg = json.load(open(CONFIG))
    boards_cfg = cfg["boards"]
    match = build_matcher(cfg.get("match", {}))
    wd_age = int(cfg.get("match", {}).get("max_age_days", 45))
    wd_cats = tuple(k.lower() for k in
                    cfg.get("match", {}).get("workday_category_keywords", ["product"]))

    raw_by_uid = {}    # everything seen on a board, before filtering
    jobs = []
    for tok, name in boards_cfg.get("greenhouse", {}).items():
        jobs.append(("Greenhouse", tok, name, lambda t=tok, n=name: fetch_greenhouse(t, n)))
    for tok, name in boards_cfg.get("lever", {}).items():
        jobs.append(("Lever", tok, name, lambda t=tok, n=name: fetch_lever(t, n)))
    for tok, name in boards_cfg.get("ashby", {}).items():
        jobs.append(("Ashby", tok, name, lambda t=tok, n=name: fetch_ashby(t, n)))
    for tok, meta in boards_cfg.get("workday", {}).items():
        jobs.append(("Workday", tok, meta.get("name", tok),
                     lambda t=tok, m=meta: fetch_workday(
                         t, m, min(wd_age, 30), cat_keywords=wd_cats)))

    print(f"Fetching {len(jobs)} boards…")
    results, board_rows = [], []

    def work(j):
        ats, tok, name, fn = j
        try:
            return j, fn(), None
        except Exception as e:                        # noqa: BLE001
            return j, [], f"{type(e).__name__}: {e}"[:120]

    with ThreadPoolExecutor(12) as ex:
        for (ats, tok, name, _), posts, err in ex.map(work, jobs):
            kept = []
            for p in posts:
                raw_by_uid[p["uid"]] = p
            for p in posts:
                ok, _why = match(p)
                if ok:
                    kept.append(p)
            results.extend(kept)
            board_rows.append([name, ats, tok, len(posts), len(kept),
                               err or "ok"])
            flag = "!" if err else " "
            print(f" {flag} {name:<26} {ats:<11} {len(posts):>5} live  {len(kept):>3} match"
                  + (f"  [{err}]" if err else ""))

    # de-dupe: same company + title can appear on two ATS
    seen, deduped = set(), []
    for p in sorted(results, key=lambda x: (x["company"].lower(), x["title"].lower())):
        k = (p["company"].lower(), p["title"].lower(), (p["location"] or "").lower())
        if k in seen:
            continue
        seen.add(k)
        deduped.append(p)

    state = load_state()
    known = state["postings"]
    edits = read_user_edits()
    today = TODAY.isoformat()

    live_uids = set()
    for p in deduped:
        uid = p["uid"]
        live_uids.add(uid)
        rec = known.get(uid, {})
        rec.update({
            "title": p["title"], "company": p["company"], "location": p["location"],
            "url": p["url"], "source": p["source"], "dept": p["dept"], "req": p["req"],
            "posted": p["posted"].isoformat() if p["posted"] else "",
            "last_seen": today,
        })
        rec.setdefault("first_seen", today)
        for c in TRACK:
            if uid in edits and c in edits[uid]:
                rec[c] = edits[uid][c]
            rec.setdefault(c, "")
        known[uid] = rec
        p["first_seen"] = rec["first_seen"]
        for c in TRACK:
            p[c] = rec[c]

    # anything previously live and now gone
    max_age = int(cfg.get("match", {}).get("max_age_days", 45))
    closed = []
    for uid, rec in known.items():
        if uid in live_uids:
            continue
        raw = raw_by_uid.get(uid)
        if raw is not None:
            # Still on the board. Either it aged past the freshness window --
            # worth recording, since the req is genuinely still open -- or the
            # filters changed, which is not a departure and is not reported.
            _ok, why = match(raw)
            if why != "stale":
                continue
            reason = "Aged out (>%d days)" % max_age
        else:
            reason = "Removed from board"
        # Keep whichever reason applied when it first left the Active tab.
        rec.setdefault("closed_reason", reason)
        if uid in edits:
            for c in TRACK:
                if c in edits[uid]:
                    rec[c] = edits[uid][c]
        fs, ls = rec.get("first_seen", ""), rec.get("last_seen", "")
        days = ""
        try:
            days = (dt.date.fromisoformat(ls) - dt.date.fromisoformat(fs)).days
        except Exception:                             # noqa: BLE001
            pass
        closed.append({**rec, "uid": uid, "days_live": days,
                       "first_seen": fs, "last_seen": ls,
                       "reason": rec.get("closed_reason", "")})
    closed.sort(key=lambda c: c.get("last_seen", ""), reverse=True)

    active = sorted(deduped, key=lambda p: (p["posted"] or dt.date(1970, 1, 1)), reverse=True)
    new_count = sum(1 for p in active if p["first_seen"] == today)

    stats = {
        "Last run": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Boards polled": len(jobs),
        "Boards erroring": sum(1 for b in board_rows if b[5] != "ok"),
        "Postings scanned": sum(b[3] for b in board_rows),
        "Matching (after de-dupe)": len(active),
        "New since last run": new_count,
        "Closed - removed from board":
            sum(1 for c in closed if c.get("reason", "").startswith("Removed")),
        "Closed - aged out":
            sum(1 for c in closed if c.get("reason", "").startswith("Aged")),
        "Previous run": state.get("last_run") or "first run",
    }
    board_rows.sort(key=lambda b: (-b[4], b[0].lower()))
    write_workbook(active, closed, board_rows, cfg, stats)

    state["last_run"] = dt.datetime.now().isoformat(timespec="seconds")
    state["postings"] = known
    save_state(state)

    print(f"\n  {len(active)} matching roles   {new_count} new   {len(closed)} closed")
    print(f"  → {XLSX}")


# ---------------------------------------------------------------- add-board

BOARD_PATTERNS = [
    (r"job-boards\.greenhouse\.io/([^/?#]+)",              "greenhouse"),
    (r"boards\.greenhouse\.io/([^/?#]+)",                  "greenhouse"),
    (r"jobs\.lever\.co/([^/?#]+)",                         "lever"),
    (r"jobs\.ashbyhq\.com/([^/?#]+)",                      "ashby"),
]


def cmd_add_board(args):
    url, cfg = args.url, json.load(open(CONFIG))
    name = args.name

    m = re.search(r"https?://([^.]+)\.wd(\d+)\.myworkdayjobs\.com/(?:[a-z]{2}-[A-Z]{2}/)?([^/?#]+)", url)
    if m:
        tok, dc, site = m.group(1), int(m.group(2)), m.group(3)
        try:
            got = fetch_workday(tok, {"dc": dc, "site": site, "name": name or tok}, 7, page_cap=1)
        except Exception as e:                        # noqa: BLE001
            print(f"Could not reach that Workday board: {e}"); return 1
        cfg["boards"].setdefault("workday", {})[tok] = {
            "name": name or tok.title(), "dc": dc, "site": site}
        json.dump(cfg, open(CONFIG, "w"), indent=2)
        print(f"Added Workday board {tok} (wd{dc}/{site}) — {len(got)} recent postings.")
        return 0

    for pat, ats in BOARD_PATTERNS:
        m = re.search(pat, url)
        if not m:
            continue
        tok = m.group(1)
        fn = {"greenhouse": fetch_greenhouse, "lever": fetch_lever, "ashby": fetch_ashby}[ats]
        try:
            got = fn(tok, name or tok)
        except Exception as e:                        # noqa: BLE001
            print(f"Could not reach that {ats} board: {e}"); return 1
        cfg["boards"].setdefault(ats, {})[tok] = name or tok.title()
        json.dump(cfg, open(CONFIG, "w"), indent=2)
        print(f"Added {ats} board {tok} — {len(got)} live postings.")
        return 0

    print("Unrecognised board URL. Expected a greenhouse / lever / ashby / "
          "myworkdayjobs.com careers URL.")
    return 1


def cmd_boards(args):
    cfg = json.load(open(CONFIG))
    for ats, boards in cfg["boards"].items():
        print(f"\n{ats} ({len(boards)})")
        for tok, v in sorted(boards.items()):
            label = v if isinstance(v, str) else f"{v.get('name')}  wd{v.get('dc')}/{v.get('site')}"
            print(f"  {tok:<22} {label}")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd")
    sub.add_parser("run").set_defaults(fn=cmd_run)
    a = sub.add_parser("add-board"); a.add_argument("url"); a.add_argument("--name", default="")
    a.set_defaults(fn=cmd_add_board)
    sub.add_parser("boards").set_defaults(fn=cmd_boards)
    args = ap.parse_args()
    if not getattr(args, "fn", None):
        args.fn, args.cmd = cmd_run, "run"
    return args.fn(args) or 0


if __name__ == "__main__":
    sys.exit(main())
