import argparse
import csv
import json
import re
import time
import urllib.parse
import urllib.request
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "journals.yml"
HISTORY_PATH = ROOT / "data" / "doi_history.csv"
MASTER_PATH = ROOT / "data" / "all_articles.xlsx"
REPORT_DIR = ROOT / "reports" / "monthly"
OUTPUT_DIR = ROOT / "output"

FIELDS = [
    "Title",
    "Publication Year",
    "Publication Date",
    "Journal",
    "DOI",
    "Citations (OpenAlex)",
    "Citation Retrieved",
    "Authors",
    "Primary Topic",
    "Screening Decision",
    "Screening Rationale",
    "Abstract",
    "Keywords / OpenAlex Topics",
    "Article URL",
    "OpenAlex ID",
]

BIOMED_PATTERNS = [
    r"\bbiostat", r"\bmedical\b", r"\bmedicine\b", r"\bhealth\b", r"\bclinical\b",
    r"\bpatient", r"\bdisease", r"\bcancer", r"\btumou?r", r"\bepidem", r"\bgenet",
    r"\bgenom", r"\btranscript", r"\bproteom", r"\bmetabol", r"\bomic", r"\bbiolog",
    r"\bbiomed", r"\bpharma", r"\bdrug\b", r"\btreatment", r"\btherapy", r"\bvaccine",
    r"\binfect", r"\bmortality", r"\bmorbidity", r"\bhospital", r"\bdiagnos",
    r"\bprognos", r"\bbrain\b", r"\bneuro", r"\bAlzheimer", r"\bcardio", r"\brenal",
    r"\bHIV\b", r"\bCOVID", r"\bSARS", r"\bpublic health", r"\bpopulation health",
]

DIRECT_METHOD_PATTERNS = [
    r"\bcausal", r"\btreatment effect", r"\bmediation", r"\binterference\b",
    r"\bsurvival\b", r"\btime[- ]to[- ]event", r"\bcensor", r"\bcompeting risk",
    r"\brecurrent event", r"\bhazard\b", r"\blongitudinal", r"\brepeated measures",
    r"\bjoint model", r"\bclinical trial", r"\brandomi[sz]ed trial", r"\badaptive trial",
    r"\bmissing data", r"\bmeasurement error", r"\bmisclassif", r"\bdiagnostic",
    r"\bROC\b", r"\bprecision medicine", r"\bdynamic treatment", r"\bpersonalized",
    r"\bprincipal stratif", r"\bcase.control", r"\bmeta-analysis", r"\bnetwork meta",
    r"\btarget trial", r"\bMendelian random", r"\bbiomarker", r"\bvaccine effect",
]

BORDERLINE_PATTERNS = [
    r"\bBayes", r"\bmachine learning", r"\bdeep learning", r"\bhigh.dimension",
    r"\bvariable selection", r"\bfunctional data", r"\bnonparametric",
    r"\bsemiparametric", r"\bprediction", r"\bclassification", r"\bspatial",
    r"\bnetwork", r"\bmultiple testing", r"\bfalse discovery",
]

TOPICS = [
    ("Causal inference and treatment effects", [
        r"\bcausal", r"\btreatment effect", r"\bmediation", r"\binterference\b",
        r"\bprincipal stratif", r"\btarget trial", r"\binstrumental variable",
        r"\bpropensity", r"\bdoubly robust", r"\btransportability", r"\bgeneraliz",
    ]),
    ("Survival and event-history analysis", [
        r"\bsurvival\b", r"\btime[- ]to[- ]event", r"\bcensor", r"\bhazard\b",
        r"\bcompeting risk", r"\brecurrent event", r"\bmultistate", r"\bcure model",
        r"\blifetime data", r"\bjoint model",
    ]),
    ("Clinical trials and experimental design", [
        r"\bclinical trial", r"\brandomi[sz](?:ed|ation)\s+(?:clinical\s+)?trial",
        r"\badaptive design", r"\btrial design", r"\bdose.find", r"\bphase [ivx0-9]+",
        r"\bnoninferiority", r"\bsequential", r"\bplatform trial", r"\bcluster trial",
        r"\bstepped wedge",
    ]),
    ("Machine learning and precision health", [
        r"\bmachine learning", r"\bdeep learning", r"\bneural network", r"\brandom forest",
        r"\bprecision medicine", r"\bpersonalized", r"\bindividuali[sz]ed treatment",
        r"\bdynamic treatment", r"\bheterogeneous treatment", r"\bprediction",
        r"\bhigh.dimension", r"\bvariable selection", r"\breinforcement learning",
    ]),
    ("Genomics, genetics and other omics", [
        r"\bgenet", r"\bgenom", r"\btranscript", r"\bproteom", r"\bmetabol",
        r"\bomic", r"\bGWAS\b", r"\bgene\b", r"\bDNA\b", r"\bRNA\b", r"\bsingle.cell",
    ]),
    ("Longitudinal and functional data", [
        r"\blongitudinal", r"\brepeated measures", r"\bfunctional data", r"\btrajectory",
        r"\bmixed model", r"\btime.varying", r"\bintensive longitudinal",
    ]),
    ("Missing data, measurement error and data fusion", [
        r"\bmissing data", r"\bimputation", r"\bmeasurement error", r"\bmisclassif",
        r"\bdata fusion", r"\bdata integration", r"\brecord linkage", r"\btwo.sample",
        r"\bvalidation sample", r"\bselection bias",
    ]),
    ("Epidemiology, infectious disease and population health", [
        r"\bepidem", r"\binfect", r"\bCOVID", r"\bSARS", r"\bvaccine", r"\bpublic health",
        r"\bpopulation health", r"\bdisease mapping", r"\boutbreak", r"\btransmission",
    ]),
    ("Diagnostic, prognostic and risk prediction", [
        r"\bdiagnos", r"\bprognos", r"\brisk prediction", r"\bROC\b", r"\bAUC\b",
        r"\bcalibration", r"\bclinical prediction", r"\bbiomarker",
    ]),
    ("Bayesian biomedical methods", [
        r"\bBayes", r"\bposterior", r"\bprior\b", r"\bMCMC\b",
    ]),
]


def api_json(url, attempts=5):
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "biostat-journal-monitor/1.0"},
    )
    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                return json.load(response)
        except Exception:
            if attempt == attempts - 1:
                raise
            time.sleep(2 ** attempt)


def reconstruct_abstract(index):
    if not index:
        return ""
    positions = []
    for word, locations in index.items():
        positions.extend((position, word) for position in locations)
    return " ".join(word for _, word in sorted(positions))


def normalize_text(value):
    return re.sub(r"\s+", " ", value or "").strip()


def matches_any(text, patterns):
    return any(re.search(pattern, text, flags=re.I) for pattern in patterns)


def load_journals(path=CONFIG_PATH):
    with path.open(encoding="utf-8") as handle:
        return yaml.safe_load(handle)["journals"]


def source_id(issn):
    data = api_json(f"https://api.openalex.org/sources/issn:{issn}")
    return data["id"].rsplit("/", 1)[-1]


def fetch_works(journal, start_date, end_date):
    cursor = "*"
    results = []
    source = source_id(journal["issn"])
    filters = (
        f"primary_location.source.id:{source},"
        f"from_publication_date:{start_date},to_publication_date:{end_date},type:article"
    )
    while cursor:
        params = urllib.parse.urlencode({
            "filter": filters,
            "per-page": 200,
            "cursor": cursor,
            "select": (
                "id,doi,title,publication_year,publication_date,authorships,"
                "cited_by_count,abstract_inverted_index,primary_location,keywords,topics"
            ),
        })
        data = api_json(f"https://api.openalex.org/works?{params}")
        batch = data["results"]
        for work in batch:
            work["_journal"] = journal["name"]
            work["_screening"] = journal["screening"]
        results.extend(batch)
        if not batch:
            break
        cursor = data["meta"].get("next_cursor")
        time.sleep(0.1)
    return results


def screen_paper(mode, title, abstract, keywords, concepts):
    text = " ".join([title, abstract, keywords, concepts])
    if mode == "core":
        return "Included", "Core biostatistics or medical-statistics journal"
    if matches_any(text, BIOMED_PATTERNS):
        return "Included", "Explicit biomedical, health, biological, or pharmaceutical focus"
    if matches_any(text, DIRECT_METHOD_PATTERNS):
        return "Included", "Method directly central to biostatistical research"
    if matches_any(text, BORDERLINE_PATTERNS):
        return "Needs Review", "Adjacent statistical method without an explicit biomedical focus"
    return "Excluded", "No direct biostatistical or biomedical focus identified"


def classify_topic(title, abstract, keywords, concepts):
    supporting = " ".join([abstract, keywords, concepts])
    best_score = 0
    best_topic = "Other biostatistical methodology"
    for topic, patterns in TOPICS:
        title_score = sum(bool(re.search(pattern, title, flags=re.I)) for pattern in patterns)
        support_score = sum(bool(re.search(pattern, supporting, flags=re.I)) for pattern in patterns)
        score = 4 * title_score + support_score
        if score > best_score:
            best_score = score
            best_topic = topic
    return best_topic


def paper_row(work, retrieved):
    title = normalize_text(work.get("title"))
    abstract = normalize_text(reconstruct_abstract(work.get("abstract_inverted_index")))
    keywords = "; ".join(
        normalize_text(item.get("display_name"))
        for item in work.get("keywords") or []
        if item.get("display_name")
    )
    concepts = "; ".join(
        normalize_text(item.get("display_name"))
        for item in work.get("topics") or []
        if item.get("display_name")
    )
    decision, rationale = screen_paper(
        work["_screening"], title, abstract, keywords, concepts
    )
    authors = "; ".join(
        normalize_text((item.get("author") or {}).get("display_name"))
        for item in work.get("authorships") or []
        if (item.get("author") or {}).get("display_name")
    )
    doi = normalize_text((work.get("doi") or "").replace("https://doi.org/", "")).lower()
    location = work.get("primary_location") or {}
    return {
        "Title": title,
        "Publication Year": str(work.get("publication_year") or ""),
        "Publication Date": work.get("publication_date") or "",
        "Journal": work["_journal"],
        "DOI": doi,
        "Citations (OpenAlex)": int(work.get("cited_by_count") or 0),
        "Citation Retrieved": retrieved,
        "Authors": authors,
        "Primary Topic": classify_topic(title, abstract, keywords, concepts),
        "Screening Decision": decision,
        "Screening Rationale": rationale,
        "Abstract": abstract,
        "Keywords / OpenAlex Topics": "; ".join(filter(None, [keywords, concepts])),
        "Article URL": f"https://doi.org/{doi}" if doi else normalize_text(location.get("landing_page_url")),
        "OpenAlex ID": work.get("id", ""),
    }


def read_history(path=HISTORY_PATH):
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8") as handle:
        return {row["doi"]: row for row in csv.DictReader(handle) if row.get("doi")}


def write_history(history, path=HISTORY_PATH):
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = ["doi", "first_seen", "publication_date", "journal", "decision", "title"]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(sorted(history.values(), key=lambda row: row["doi"]))


def deduplicate(rows):
    unique = {}
    for row in sorted(rows, key=lambda item: item["Publication Date"], reverse=True):
        key = row["DOI"] or row["OpenAlex ID"]
        unique.setdefault(key, row)
    return list(unique.values())


def new_rows(rows, history):
    return [row for row in rows if row["DOI"] and row["DOI"] not in history]


def update_history(history, rows, first_seen):
    for row in rows:
        doi = row["DOI"]
        if doi and doi not in history:
            history[doi] = {
                "doi": doi,
                "first_seen": first_seen,
                "publication_date": row["Publication Date"],
                "journal": row["Journal"],
                "decision": row["Screening Decision"],
                "title": row["Title"],
            }


def write_sheet(sheet, rows, fields=FIELDS):
    sheet.append(fields)
    for row in rows:
        sheet.append([row.get(field, "") for field in fields])


def style_workbook(workbook):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10)
    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A2"
        if sheet.max_row > 1:
            sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 30
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        widths = [55, 12, 15, 44, 28, 18, 18, 50, 38, 18, 48, 70, 60, 38, 28]
        for index, width in enumerate(widths[:sheet.max_column], start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for row_index in range(2, sheet.max_row + 1):
            sheet.row_dimensions[row_index].height = 35


def create_report_workbook(rows, path, start_date, end_date, retrieved):
    workbook = Workbook()
    workbook.remove(workbook.active)
    decisions = ["Included", "Needs Review", "Excluded"]
    for decision in decisions:
        sheet = workbook.create_sheet(decision.replace(" ", "_"))
        write_sheet(sheet, [row for row in rows if row["Screening Decision"] == decision])
    summary = workbook.create_sheet("Summary", 0)
    summary.append(["Metric", "Value"])
    summary.append(["Search start", start_date])
    summary.append(["Search end", end_date])
    summary.append(["Retrieved", retrieved])
    for decision in decisions:
        summary.append([
            decision,
            sum(row["Screening Decision"] == decision for row in rows),
        ])
    for journal, count in Counter(row["Journal"] for row in rows).most_common():
        summary.append([f"Journal: {journal}", count])
    style_workbook(workbook)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def load_master(path=MASTER_PATH):
    if not path.exists():
        return []
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["Articles"]
    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    rows = [
        dict(zip(headers, values))
        for values in sheet.iter_rows(min_row=2, values_only=True)
    ]
    workbook.close()
    return rows


def write_master(rows, path=MASTER_PATH):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Articles"
    write_sheet(sheet, rows)
    style_workbook(workbook)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def merge_master(existing, additions):
    combined = {}
    for row in existing + additions:
        key = row.get("DOI") or row.get("OpenAlex ID")
        if key:
            combined[key] = {field: row.get(field, "") for field in FIELDS}
    return sorted(
        combined.values(),
        key=lambda row: (row["Publication Date"], row["Journal"], row["Title"]),
        reverse=True,
    )


def write_markdown(rows, path, start_date, end_date, retrieved):
    included = [row for row in rows if row["Screening Decision"] == "Included"]
    review = [row for row in rows if row["Screening Decision"] == "Needs Review"]
    topic_counts = Counter(row["Primary Topic"] for row in included)
    journal_counts = Counter(row["Journal"] for row in included)
    lines = [
        f"# Biostatistics journal update: {end_date}",
        "",
        f"Search window: {start_date} through {end_date}. Retrieved: {retrieved}.",
        "",
        f"New included papers: {len(included)}",
        f"New papers needing review: {len(review)}",
        "",
        "## Included papers by journal",
        "",
    ]
    for journal, count in journal_counts.most_common():
        lines.append(f"- {journal}: {count}")
    lines.extend(["", "## Included papers by topic", ""])
    for topic, count in topic_counts.most_common():
        lines.append(f"- {topic}: {count}")
    lines.extend(["", "## Included papers", ""])
    for row in sorted(included, key=lambda item: item["Publication Date"], reverse=True):
        lines.append(
            f"- [{row['Title']}]({row['Article URL']}) — {row['Journal']} "
            f"({row['Publication Date']}); {row['Primary Topic']}."
        )
    if review:
        lines.extend(["", "## Needs review", ""])
        for row in sorted(review, key=lambda item: item["Publication Date"], reverse=True):
            lines.append(
                f"- [{row['Title']}]({row['Article URL']}) — {row['Journal']} "
                f"({row['Publication Date']})."
            )
    lines.extend([
        "",
        "Citation counts are OpenAlex cited-by counts and are time-sensitive.",
    ])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def write_run_output(rows, report_stem, start_date, end_date):
    included = sum(row["Screening Decision"] == "Included" for row in rows)
    review = sum(row["Screening Decision"] == "Needs Review" for row in rows)
    issue_path = OUTPUT_DIR / "issue_body.md"
    issue_lines = [
        f"Search window: {start_date} through {end_date}.",
        "",
        f"- New included papers: {included}",
        f"- New papers needing review: {review}",
        f"- [Monthly Markdown report](../blob/main/reports/monthly/{report_stem}.md)",
        f"- [Monthly Excel report](../blob/main/reports/monthly/{report_stem}.xlsx)",
        f"- [Cumulative workbook](../blob/main/data/all_articles.xlsx)",
    ]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    issue_path.write_text("\n".join(issue_lines), encoding="utf-8")
    summary = {
        "new_count": len(rows),
        "included_count": included,
        "review_count": review,
        "report_stem": report_stem,
        "issue_title": f"Biostatistics journal update — {end_date} ({included} included)",
    }
    (OUTPUT_DIR / "run_summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    return summary


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--lookback-days", type=int, default=45)
    parser.add_argument("--as-of", help="End date in YYYY-MM-DD format")
    return parser.parse_args()


def main():
    args = parse_args()
    end = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else date.today()
    start = end - timedelta(days=args.lookback_days - 1)
    retrieved = date.today().isoformat()
    works = []
    for journal in load_journals():
        works.extend(fetch_works(journal, start.isoformat(), end.isoformat()))
    rows = deduplicate([paper_row(work, retrieved) for work in works])
    history = read_history()
    unseen = new_rows(rows, history)
    report_stem = f"{end:%Y-%m-%d}_lookback-{args.lookback_days}d"
    if unseen:
        report_xlsx = REPORT_DIR / f"{report_stem}.xlsx"
        report_md = REPORT_DIR / f"{report_stem}.md"
        create_report_workbook(
            unseen, report_xlsx, start.isoformat(), end.isoformat(), retrieved
        )
        write_markdown(unseen, report_md, start.isoformat(), end.isoformat(), retrieved)
        master_additions = [
            row for row in unseen
            if row["Screening Decision"] in {"Included", "Needs Review"}
        ]
        if master_additions:
            write_master(merge_master(load_master(), master_additions))
        update_history(history, unseen, retrieved)
        write_history(history)
    summary = write_run_output(
        unseen, report_stem, start.isoformat(), end.isoformat()
    )
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
