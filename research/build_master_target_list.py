#!/usr/bin/env python3
"""
Consolidate Kalpesh Kinariwala's LinkedIn target spreadsheets into one
clean, deduplicated master target list.

Sources (all read from the upload folder; NO web scraping):
  A. Kalpesh_UAE_Target_Accounts.xlsx  -> sheet 'Kalpesh Target List'
  B. Kalpesh_LinkedIn_Strategy_Sheet.xlsx -> sheet 'Targeted LinkedIn Accounts'
  C. LinkedIn_200_Accounts.xlsx -> sheet 'LinkedIn Accounts'

The other supplied files (KPI tracker, master execution sheet) and the
secondary tabs (Read Me, Profile & Strategy, Niche Summary, Summary by Niche)
contain no people records and are intentionally not loaded.

Output: KK_Master_Target_List.xlsx  (Master List + Data Quality Log)
"""
import os, re, unicodedata
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

UP = "/root/.claude/uploads/d502465e-97b0-40ad-b7b2-4e48e45e93b6"
OUT = "/home/user/kk/KK_Master_Target_List.xlsx"

SRC = {
    "UAE_Target_Accounts": ("32ba5d00-Kalpesh_UAE_Target_Accounts.xlsx", "Kalpesh Target List", 1),
    "LinkedIn_Strategy":   ("383d4ace-Kalpesh_LinkedIn_Strategy_Sheet.xlsx", "Targeted LinkedIn Accounts", 2),
    "LinkedIn_200":        ("d8ab0841-LinkedIn_200_Accounts.xlsx", "LinkedIn Accounts", 0),
}

# ---------------------------------------------------------------- helpers
URL_PLACEHOLDERS = ("linkedin search", "no direct profile", "verify", "n/a", "tbd")

def s(x):
    if x is None:
        return ""
    x = str(x).strip()
    return "" if x.lower() in ("nan", "none") else x

def is_placeholder_name(name):
    n = s(name)
    return (n == "") or ("verify" in n.lower())

def strip_accents(t):
    return "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))

def norm_name(name):
    if is_placeholder_name(name):
        return ""
    n = strip_accents(s(name)).lower()
    n = re.sub(r"[^a-z0-9 ]", " ", n)
    return re.sub(r"\s+", " ", n).strip()

def is_real_url(raw):
    # A real URL is an actual linkedin path. Placeholders ("VERIFY",
    # "LinkedIn search", "No direct profile", "n/a", "tbd", "") never contain
    # "linkedin.com/", so this positive test cleanly excludes them. (A naive
    # substring scan would wrongly flag handles like ".../in/andrew" — the
    # "in/a" contains "n/a".)
    return "linkedin.com/" in s(raw).lower()

def norm_url(raw):
    """Canonical comparison/display form: strip protocol, www, ae./sa., trailing slash."""
    if not is_real_url(raw):
        return ""
    u = s(raw).lower()
    u = re.sub(r"^https?://", "", u)
    u = re.sub(r"^www\.", "", u)
    u = re.sub(r"^(ae|sa)\.", "", u)
    return u.rstrip("/")

def geography(loc):
    loc = s(loc)
    if loc == "" or loc.lower().startswith("verify"):
        return "Other"
    token = loc.split(",")[-1].strip().upper() if "," in loc else loc.split("/")[0].strip().upper()
    if "UAE" in token:
        return "UAE"
    if token in ("UK", "GB", "ENGLAND", "U.K."):
        return "UK"
    if token in ("US", "USA", "U.S."):
        return "US"
    return "Other"

def clean_title(t):
    t = s(t)
    t = re.sub(r"\s*\(confirm[^)]*\)", "", t, flags=re.I).strip()
    return t

def clean_company(c):
    c = s(c)
    if "verify" in c.lower():
        return ""
    return c

RE_BUCKETS = {"Real Estate Dev", "RE Agency", "PropTech/AI"}

def sector_bucket(industry, subniche, company, title, source):
    ind = s(industry).lower()
    txt = " ".join(s(x).lower() for x in (industry, subniche, company, title))

    # Live entertainment / artists / events
    if any(k in ind for k in ("artist management", "concerts", "live entertainment")) or \
       any(k in txt for k in ("artist manage", "concert", "live nation", "aeg presents",
                              " touring", "music week", "entertainment")):
        return "Live Entertainment"

    # Property technology (and AI applied to property)
    if "proptech" in ind:
        return "PropTech/AI"

    # Explicit real-estate dev / agency tags
    if "real estate development" in ind:
        return "Real Estate Dev"
    if "real estate agency" in ind:
        return "RE Agency"

    # Generic "Real Estate" (sheet C) or sheet-A developers
    if ind == "real estate":
        if source == "UAE_Target_Accounts":
            return "Real Estate Dev"            # sheet A is all UAE developers
        comp = s(company).lower()
        if any(k in comp for k in ("rightmove", "costar", "zillow", "portal")):
            return "PropTech/AI"
        if any(k in comp for k in ("british land", "londonmetric", "prologis", "simon property")):
            return "Real Estate Dev"
        if any(k in comp for k in ("moneyfacts", "landbay")):
            return "Investor/Finance"
        return "RE Agency"                      # remaining generic RE = agents/brokers

    # Investors / finance / capital
    if any(k in ind for k in ("investor", "finance", "vc / startup", "vc/startup")):
        return "Investor/Finance"

    # Marketing
    if any(k in ind for k in ("digital marketing", "marketing", "agency")):
        return "Marketing"

    return "Other"

def relevance(bucket, geo):
    if bucket in RE_BUCKETS:
        return "High" if geo == "UAE" else "Medium"
    if geo == "UAE" and bucket in ("Investor/Finance", "Live Entertainment"):
        return "Medium"
    return "Low"

def gen_why(bucket, geo, rel, title):
    if rel == "High":
        return f"UAE {bucket} — prime peer/target in KK's home market"
    if rel == "Medium":
        if bucket in RE_BUCKETS:
            return f"{bucket} outside UAE ({geo}) — RE peer / benchmark"
        if bucket == "Investor/Finance":
            return "UAE investor / capital audience — RE-adjacent"
        if bucket == "Live Entertainment":
            return "UAE events & leisure — Pantheon event-partner angle"
    extra = f", {geo}" if geo != "UAE" else ""
    return f"Off-sector for KK ({bucket}{extra})"

def short(t, n=95):
    t = s(t)
    return t if len(t) <= n else t[: n - 1].rstrip() + "…"

# ---------------------------------------------------------------- load
records = []   # raw, one per source row

def load(label, fn, sheet, hdr):
    df = pd.read_excel(os.path.join(UP, fn), sheet_name=sheet, header=hdr, dtype=str)
    cols = {c: c for c in df.columns}
    def col(*names):
        for n in names:
            if n in cols:
                return n
        return None
    c_name = col("Full Name")
    c_url  = col("LinkedIn URL")
    c_comp = col("Company / Business")
    c_loc  = col("Location", "HQ Location")
    c_ind  = col("Industry / Niche", "Niche")
    c_sub  = col("Sub-Niche")
    c_foll = col("Est. Followers")
    c_type = col("Account Type")
    c_why  = col("Why Relevant to KK", "Best For (Kalpesh)", "Best For (Agency Use)")
    n = 0
    for _, r in df.iterrows():
        name_raw = s(r.get(c_name))
        url_raw  = s(r.get(c_url))
        comp_raw = s(r.get(c_comp))
        # skip junk / empty trailing rows: need at least a real name, real url, or a company
        if is_placeholder_name(name_raw) and not is_real_url(url_raw) and clean_company(comp_raw) == "":
            continue
        records.append({
            "source": label,
            "name_raw": name_raw,
            "url_raw": url_raw,
            "company_raw": comp_raw,
            "loc": s(r.get(c_loc)),
            "industry": s(r.get(c_ind)),
            "subniche": s(r.get(c_sub)),
            "followers": s(r.get(c_foll)),
            "title": s(r.get(c_type)),
            "why_src": s(r.get(c_why)),
        })
        n += 1
    return n

for label, (fn, sheet, hdr) in SRC.items():
    cnt = load(label, fn, sheet, hdr)
    print(f"loaded {label}: {cnt} people rows")

# Follower figures that are placeholders should not be carried as figures
for rec in records:
    if "verify" in rec["followers"].lower():
        rec["followers"] = ""

# ---------------------------------------------------------------- union-find merge
parent = list(range(len(records)))
def find(i):
    while parent[i] != i:
        parent[i] = parent[parent[i]]
        i = parent[i]
    return i
def union(i, j):
    ri, rj = find(i), find(j)
    if ri != rj:
        parent[max(ri, rj)] = min(ri, rj)

by_name, by_url = {}, {}
for i, rec in enumerate(records):
    nn, nu = norm_name(rec["name_raw"]), norm_url(rec["url_raw"])
    if nn:
        by_name.setdefault(nn, []).append(i)
    if nu:
        by_url.setdefault(nu, []).append(i)
for idxs in list(by_name.values()) + list(by_url.values()):
    for k in range(1, len(idxs)):
        union(idxs[0], idxs[k])

# Targeted company-merges: sheet-A anonymous (name=VERIFY) rows whose company maps
# to exactly one named person in the strategy sheet.
COMPANY_MERGE = {"samana developers": "samana", "huspy": "huspy"}
def comp_key(c):
    return re.sub(r"\s+", " ", clean_company(c).lower()).strip()
named_by_company = {}
for i, rec in enumerate(records):
    if not is_placeholder_name(rec["name_raw"]) and is_real_url(rec["url_raw"]):
        named_by_company.setdefault(comp_key(rec["company_raw"]), []).append(i)
for i, rec in enumerate(records):
    if is_placeholder_name(rec["name_raw"]):
        ck = comp_key(rec["company_raw"])
        cands = named_by_company.get(ck, [])
        if len(cands) == 1:
            union(i, cands[0])

clusters = {}
for i in range(len(records)):
    clusters.setdefault(find(i), []).append(i)

# ---------------------------------------------------------------- build merged rows
SRC_NICE = {"UAE_Target_Accounts": "UAE_Target_Accounts",
            "LinkedIn_Strategy": "LinkedIn_Strategy", "LinkedIn_200": "LinkedIn_200"}
dq = []   # data-quality log rows: dict(category, person, field, kept, discarded, note)

people = []
for root, idxs in clusters.items():
    recs = [records[i] for i in idxs]
    # authoritative record = one with a real url (prefer Strategy sheet), else first
    real = [r for r in recs if is_real_url(r["url_raw"])]
    real.sort(key=lambda r: 0 if r["source"] == "LinkedIn_Strategy" else 1)
    auth = real[0] if real else recs[0]

    # ---- name
    real_names = []
    for r in recs:
        if not is_placeholder_name(r["name_raw"]):
            real_names.append(r["name_raw"])
    if real:
        name = next((r["name_raw"] for r in real if not is_placeholder_name(r["name_raw"])),
                    real_names[0] if real_names else "")
    else:
        name = real_names[0] if real_names else ""
    name_needs_find = (name == "")
    # log conflicting name spellings
    distinct_names = []
    for rn in real_names:
        if rn not in distinct_names:
            distinct_names.append(rn)
    if len(distinct_names) > 1:
        dq.append(dict(category="Conflict resolved", person=name,
                       field="Full Name", kept=name,
                       discarded="; ".join(n for n in distinct_names if n != name),
                       note="Same person, differing spelling/case; kept the variant tied to the real profile URL."))

    # ---- url
    url_norm = norm_url(auth["url_raw"]) if is_real_url(auth["url_raw"]) else ""
    if not url_norm:
        for r in real:
            if norm_url(r["url_raw"]):
                url_norm = norm_url(r["url_raw"]); break
    profile_verified = "Y" if url_norm else "N"
    placeholder_seen = any((not is_real_url(r["url_raw"])) and s(r["url_raw"]) != "" for r in recs)
    if url_norm and placeholder_seen:
        dq.append(dict(category="Conflict resolved", person=name or auth["company_raw"],
                       field="LinkedIn_URL", kept=url_norm,
                       discarded="VERIFY / placeholder",
                       note="One source had a real profile URL, another a placeholder; kept the real URL."))
    final_url = url_norm if url_norm else "NEEDS MANUAL FIND"

    # ---- company (most complete, non-placeholder)
    comps = [clean_company(r["company_raw"]) for r in recs if clean_company(r["company_raw"])]
    company = max(comps, key=len) if comps else "NEEDS MANUAL FIND"
    distinct_comps = []
    for c in comps:
        if c not in distinct_comps:
            distinct_comps.append(c)
    if len(distinct_comps) > 1:
        dq.append(dict(category="Conflict resolved", person=name or company,
                       field="Company", kept=company,
                       discarded="; ".join(c for c in distinct_comps if c != company),
                       note="Kept the most complete / specific company name."))

    # ---- title (prefer authoritative, else longest)
    titles = [clean_title(r["title"]) for r in recs if clean_title(r["title"])]
    title = clean_title(auth["title"]) or (max(titles, key=len) if titles else "")

    # ---- industry/subniche (most specific = longest non-empty)
    inds = [r["industry"] for r in recs if s(r["industry"])]
    industry = max(inds, key=len) if inds else ""
    subs = [r["subniche"] for r in recs if s(r["subniche"])]
    subniche = max(subs, key=len) if subs else ""

    # ---- followers (first real figure; log differing real figures)
    foll_vals = [r["followers"] for r in recs if s(r["followers"])]
    distinct_foll = []
    for f in foll_vals:
        if f not in distinct_foll:
            distinct_foll.append(f)
    est_followers = distinct_foll[0] if distinct_foll else ""
    if len(distinct_foll) > 1:
        dq.append(dict(category="Conflict resolved", person=name or company,
                       field="Est_Followers_UNVERIFIED", kept=est_followers,
                       discarded="; ".join(distinct_foll[1:]),
                       note="Differing UNVERIFIED follower estimates across sources; kept first. All require manual verification."))

    # ---- geography: from the authoritative (real-URL) record; fall back to
    # any record that yields a non-Other classification.
    geo = geography(auth["loc"])
    if geo == "Other":
        for r in recs:
            g = geography(r["loc"])
            if g != "Other":
                geo = g
                break

    bucket = sector_bucket(industry, subniche, company, title, auth["source"])
    rel = relevance(bucket, geo)

    # ---- why relevant: prefer curated note, else generated
    why_notes = {r["source"]: s(r["why_src"]) for r in recs if s(r["why_src"])}
    why = why_notes.get("LinkedIn_Strategy") or why_notes.get("UAE_Target_Accounts") or ""
    if not why and rel != "Low":
        why = why_notes.get("LinkedIn_200", "")
    if not why:
        why = gen_why(bucket, geo, rel, title)
    why = short(why)

    sources = sorted({SRC_NICE[r["source"]] for r in recs},
                     key=lambda x: ["UAE_Target_Accounts", "LinkedIn_Strategy", "LinkedIn_200"].index(x))

    if name_needs_find:
        display_name = "NEEDS MANUAL FIND (founder/CEO)"
        dq.append(dict(category="Needs manual NAME", person=f"[{company}]",
                       field="Full Name", kept="(unknown)", discarded="",
                       note=f"Founder/CEO name not in source data for {company} ({geo}); identify manually."))
    else:
        display_name = name

    if profile_verified == "N":
        dq.append(dict(category="Needs manual URL", person=display_name,
                       field="LinkedIn_URL", kept="NEEDS MANUAL FIND", discarded="",
                       note=f"Source had no real profile URL ({company}, {geo}); locate & paste LinkedIn URL."))

    people.append({
        "Full Name": display_name,
        "Title": title,
        "Company": company,
        "Sector_Bucket": bucket,
        "Geography": geo,
        "LinkedIn_URL": final_url,
        "Profile_Verified": profile_verified,
        "Est_Followers_UNVERIFIED": est_followers,
        "Verified_Follower_Count": "",
        "Date_Checked": "",
        "Relevance_to_KK": rel,
        "Why_Relevant": why,
        "Source_Files": "; ".join(sources),
        "_merged_from": len(recs),
    })
    if len(recs) > 1:
        dq.append(dict(category="Merged duplicate", person=display_name,
                       field="—", kept=f"1 row (from {len(recs)})",
                       discarded="; ".join(sorted({r['source'] for r in recs})),
                       note="Same person across sources (matched on normalized name and/or LinkedIn URL)."))

# ---------------------------------------------------------------- sort
GEO_ORDER = {"UAE": 0, "UK": 1, "US": 2, "Other": 3}
REL_ORDER = {"High": 0, "Medium": 1, "Low": 2}
people.sort(key=lambda p: (GEO_ORDER[p["Geography"]], REL_ORDER[p["Relevance_to_KK"]], p["Company"].lower(), p["Full Name"].lower()))
for i, p in enumerate(people, 1):
    p["#"] = i

# ---------------------------------------------------------------- summary
from collections import Counter
n_total = len(people)
by_bucket = Counter(p["Sector_Bucket"] for p in people)
by_geo = Counter(p["Geography"] for p in people)
n_uae = by_geo.get("UAE", 0)
n_y = sum(1 for p in people if p["Profile_Verified"] == "Y")
n_n = n_total - n_y

print("\n" + "=" * 64)
print("KK MASTER TARGET LIST — SUMMARY")
print("=" * 64)
print(f"Total unique people after dedupe : {n_total}")
print("\nBy Sector_Bucket:")
for b in ["Real Estate Dev", "RE Agency", "PropTech/AI", "Live Entertainment", "Investor/Finance", "Marketing", "Other"]:
    print(f"   {b:<20} {by_bucket.get(b,0)}")
print("\nBy Geography:")
for g in ["UAE", "UK", "US", "Other"]:
    print(f"   {g:<6} {by_geo.get(g,0)}")
print(f"\nUAE-only count            : {n_uae}")
print(f"Profile_Verified = Y / N  : {n_y} / {n_n}")
print("\nPriority buckets — UAE names with verified profiles, and gap to 50:")
for b in ["Real Estate Dev", "RE Agency", "PropTech/AI"]:
    uae_b = [p for p in people if p["Sector_Bucket"] == b and p["Geography"] == "UAE"]
    uae_b_ver = [p for p in uae_b if p["Profile_Verified"] == "Y"]
    gap = max(0, 50 - len(uae_b_ver))
    print(f"   {b:<18} UAE total={len(uae_b):>2}  verified(Y)={len(uae_b_ver):>2}  gap to 50={gap}")
print("=" * 64)

# expose for the writer step
globals()["PEOPLE"] = people
globals()["DQ"] = dq
print(f"\nData-quality log entries: {len(dq)}")

# ---------------------------------------------------------------- methodology notes
n_with_foll = sum(1 for p in people if p["Est_Followers_UNVERIFIED"])
notes = [
    ("Methodology note", "Sources merged", "—",
     "UAE_Target_Accounts + LinkedIn_Strategy + LinkedIn_200",
     "KPI Tracker, Master Execution Sheet, and summary/profile tabs (no people records)",
     "The three named sheets are the only ones holding individual target records; the rest are excluded."),
    ("Methodology note", "Dedupe key", "—", "Normalized full name AND/OR LinkedIn URL", "",
     "Names lower-cased & de-accented; URLs stripped of http(s)://, www., ae./sa., trailing '/'. Rows merge if either key matches."),
    ("Methodology note", "Follower figures", "Est_Followers_UNVERIFIED",
     f"{n_with_foll} rows carry a source estimate (verbatim)", "",
     "All follower counts are UNVERIFIED source estimates. Verified_Follower_Count and Date_Checked are intentionally BLANK for manual entry. No counts were invented."),
    ("Methodology note", "VERIFY placeholders", "LinkedIn_URL / Est_Followers",
     "Set to NEEDS MANUAL FIND / left blank", "",
     "Sheet A used 'VERIFY' where data was deliberately not fabricated. Treated as missing -> Profile_Verified = N."),
    ("Methodology note", "Geography rule", "Geography", "Primary location token -> UAE/UK/US/Other", "",
     "Multi-region tags use the FIRST token: 'US/UAE'->US, 'UAE/IN'->UAE, 'IN/UAE'->Other. UAE = UAE-primary only (conservative UAE count)."),
    ("Methodology note", "Relevance rule", "Relevance_to_KK", "High / Medium / Low", "",
     "High = UAE RE Dev/Agency/PropTech. Medium = RE outside UAE, or UAE Investor/Finance or Live Entertainment (RE-adjacent). Low = off-sector."),
    ("Methodology note", "Bucket note", "Sector_Bucket", "PropTech/AI = property technology", "",
     "General AI/tech founders (e.g. NVIDIA, Cohere) are classified 'Other', not PropTech/AI."),
    ("Methodology note", "Org-not-person", "Full Name",
     "Propy; AEG Presents; Galliard Group", "",
     "A few source rows are company/brand pages rather than individuals; kept with their LinkedIn company URL."),
]
CAT_RANK = {"Methodology note": 0, "Conflict resolved": 1, "Merged duplicate": 2,
            "Needs manual NAME": 3, "Needs manual URL": 4}
def dq_row_tuple(d):
    if isinstance(d, tuple):
        return d
    return (d["category"], d["person"], d["field"], d["kept"], d["discarded"], d["note"])
dq_all = notes + [dq_row_tuple(d) for d in dq]
dq_all.sort(key=lambda t: (CAT_RANK.get(t[0], 9),))

# ---------------------------------------------------------------- write xlsx
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BASE_FONT = Font(name="Calibri", size=11)
REL_FILL = {"High": PatternFill("solid", fgColor="C6EFCE"),
            "Medium": PatternFill("solid", fgColor="FFEB9C"),
            "Low": PatternFill("solid", fgColor="F2F2F2")}
NEED_FILL = PatternFill("solid", fgColor="FFC7CE")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")

wb = openpyxl.Workbook()

# -- Master List
ws = wb.active
ws.title = "Master List"
COLS = ["#", "Full Name", "Title", "Company", "Sector_Bucket", "Geography",
        "LinkedIn_URL", "Profile_Verified", "Est_Followers_UNVERIFIED",
        "Verified_Follower_Count", "Date_Checked", "Relevance_to_KK",
        "Why_Relevant", "Source_Files"]
ws.append(COLS)
for p in people:
    ws.append([p[c] for c in COLS])

rel_idx = COLS.index("Relevance_to_KK") + 1
pv_idx = COLS.index("Profile_Verified") + 1
url_idx = COLS.index("LinkedIn_URL") + 1
name_idx = COLS.index("Full Name") + 1
for j, c in enumerate(COLS, 1):
    cell = ws.cell(row=1, column=j)
    cell.fill = HDR_FILL; cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
for i, p in enumerate(people, start=2):
    for j, c in enumerate(COLS, 1):
        cell = ws.cell(row=i, column=j)
        cell.font = BASE_FONT
        cell.border = BORDER
        wrap = c in ("Why_Relevant", "Company", "Title", "Source_Files")
        cell.alignment = Alignment(horizontal="center" if c in ("#", "Geography", "Profile_Verified", "Relevance_to_KK") else "left",
                                   vertical="top", wrap_text=wrap)
    ws.cell(row=i, column=rel_idx).fill = REL_FILL[p["Relevance_to_KK"]]
    if p["Profile_Verified"] == "N":
        ws.cell(row=i, column=pv_idx).fill = NEED_FILL
    if p["LinkedIn_URL"] == "NEEDS MANUAL FIND":
        ws.cell(row=i, column=url_idx).fill = NEED_FILL
    if str(p["Full Name"]).startswith("NEEDS MANUAL FIND"):
        ws.cell(row=i, column=name_idx).fill = NEED_FILL

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(people)+1}"

WIDTHS = {"#": 5, "Full Name": 26, "Title": 20, "Company": 26, "Sector_Bucket": 17,
          "Geography": 10, "LinkedIn_URL": 38, "Profile_Verified": 9,
          "Est_Followers_UNVERIFIED": 13, "Verified_Follower_Count": 13,
          "Date_Checked": 12, "Relevance_to_KK": 13, "Why_Relevant": 52, "Source_Files": 24}
for j, c in enumerate(COLS, 1):
    ws.column_dimensions[get_column_letter(j)].width = WIDTHS[c]
ws.row_dimensions[1].height = 30

# -- Data Quality Log
ws2 = wb.create_sheet("Data Quality Log")
DCOLS = ["Category", "Person / Company", "Field", "Value Kept",
         "Value(s) Discarded / Missing", "Reason / Note"]
ws2.append(DCOLS)
for t in dq_all:
    ws2.append(list(t))
for j, c in enumerate(DCOLS, 1):
    cell = ws2.cell(row=1, column=j)
    cell.fill = HDR_FILL; cell.font = HDR_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
CAT_FILL = {"Methodology note": PatternFill("solid", fgColor="DDEBF7"),
            "Conflict resolved": PatternFill("solid", fgColor="FFF2CC"),
            "Merged duplicate": PatternFill("solid", fgColor="E2EFDA"),
            "Needs manual NAME": PatternFill("solid", fgColor="FCE4D6"),
            "Needs manual URL": PatternFill("solid", fgColor="FCE4D6")}
for i, t in enumerate(dq_all, start=2):
    for j in range(1, len(DCOLS) + 1):
        cell = ws2.cell(row=i, column=j)
        cell.font = BASE_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="left", vertical="top",
                                   wrap_text=(j in (5, 6)))
    ws2.cell(row=i, column=1).fill = CAT_FILL.get(t[0], PatternFill())
ws2.freeze_panes = "A2"
ws2.auto_filter.ref = f"A1:{get_column_letter(len(DCOLS))}{len(dq_all)+1}"
DWIDTHS = [18, 26, 22, 34, 34, 64]
for j, w in enumerate(DWIDTHS, 1):
    ws2.column_dimensions[get_column_letter(j)].width = w
ws2.row_dimensions[1].height = 28

wb.save(OUT)
print(f"\nWrote {OUT}  ({len(people)} master rows, {len(dq_all)} data-quality rows)")
