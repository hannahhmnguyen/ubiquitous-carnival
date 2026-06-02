"""
CEO Finder — finds missing CEO/founder names for clinic companies.

For each row missing a name, runs 5 searches via Claude's web_search tool:
  1. "[Company]" CEO
  2. "[Company]" founder
  3. site:[domain] CEO OR founder OR leadership
  4. "[Company]" LinkedIn CEO
  5. "[Company]" Crunchbase CEO

Outputs: Company | CEO First | CEO Last | Title | Source URL | Confidence | Notes

Usage:
    export ANTHROPIC_API_KEY=sk-...
    python3 ceo_finder.py --input Womens_HRT_Hormones_Only_New_cleaned.xlsx
    python3 ceo_finder.py --input Womens_HRT_Hormones_Only_New_cleaned.xlsx --output ceo_results.xlsx --batch 10
"""

import argparse
import json
import os
import re
import time
from urllib.parse import urlparse

import anthropic
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

MODEL = "claude-opus-4-8"  # supports web_search tool

CONFIDENCE_COLORS = {
    "High":   "C6EFCE",  # green
    "Medium": "FFEB9C",  # yellow
    "Low":    "FFC7CE",  # red/pink
    "":       "FFFFFF",
}

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def get_companies(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Deduped"]
    rows = list(ws.iter_rows(values_only=True))
    companies = []
    for i, r in enumerate(rows[1:], 1):
        companies.append({
            "row":              i,
            "company":          str(r[0]).strip() if r[0] else "",
            "existing_first":   str(r[1]).strip() if r[1] else "",
            "existing_last":    str(r[2]).strip() if r[2] else "",
            "website":          str(r[3]).strip() if r[3] else "",
            "existing_phone":   str(r[4]).strip() if r[4] else "",
            "existing_title":   str(r[5]).strip() if r[5] else "",
            "existing_email":   str(r[6]).strip() if r[6] else "",
            "existing_linkedin":str(r[7]).strip() if r[7] else "",
            "reviews":          r[8],
        })
    return companies


def normalize_url(website):
    if not website:
        return None
    w = website.strip()
    if not w.startswith("http"):
        w = "https://" + w
    return w.rstrip("/")


def domain_of(url):
    if not url:
        return ""
    try:
        return urlparse(url).netloc.lstrip("www.")
    except Exception:
        return ""

# ---------------------------------------------------------------------------
# Name-in-company-field detection
# Catches rows like "Dr. Reem Sharhan, ND" or "...Pamela Gaudry, M.D."
# ---------------------------------------------------------------------------

INLINE_NAME_PATTERNS = [
    # "Name: Pamela Gaudry, M.D." or "...: First Last, M.D."
    re.compile(r":\s*([A-Z][a-z]+(?:\s[A-Z][a-z]+)+),?\s*(?:M\.?D\.?|D\.?O\.?|N\.?D\.?|Ph\.?D\.?|APRN|NP|PA)", re.IGNORECASE),
    # "Dr. First Last" at start or after colon
    re.compile(r"(?:^|:\s*)Dr\.?\s+([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)"),
    # "with First Last, MD"
    re.compile(r"with\s+([A-Z][a-z]+(?:\s[A-Z][a-z]+)+),?\s*(?:M\.?D\.?|D\.?O\.?)", re.IGNORECASE),
]

def extract_name_from_company_field(company_name):
    for pat in INLINE_NAME_PATTERNS:
        m = pat.search(company_name)
        if m:
            full = m.group(1).strip()
            parts = full.split()
            if len(parts) >= 2:
                return parts[0], " ".join(parts[1:]), "Extracted from company name field"
    return None, None, None

# ---------------------------------------------------------------------------
# Claude web-search researcher
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are a business intelligence researcher. Your job is to find the CEO, \
founder, or medical director of a healthcare clinic. You will be given a \
company name and website domain.

Run the specified web searches, review the results carefully, and extract \
the most reliable name you can find. Apply these rules strictly:

CONFIDENCE LEVELS:
- High: result comes from the company's own website domain OR the person's \
  own LinkedIn profile page (linkedin.com/in/...)
- Medium: result comes from Crunchbase, Bloomberg, a news article, or a \
  local business profile that clearly identifies this specific company
- Low: result comes from an aggregator (ZoomInfo, RocketReach, Apollo, etc.) \
  OR multiple sources give conflicting names

DISAMBIGUATION RULES:
- Always confirm the found person is associated with THIS specific company, \
  not a different company with a similar name.
- Prefer sources whose domain matches the company's website domain.
- If location is known, prefer results that mention the same city/state.
- If two sources disagree on the name, mark Low and note the conflict.
- If you cannot confirm the result is for the right company, mark Low.

OUTPUT FORMAT — respond ONLY with a single JSON object, no other text:
{
  "first_name": "Jane",
  "last_name": "Smith",
  "title": "Founder & Medical Director",
  "source_url": "https://example.com/about",
  "confidence": "High",
  "notes": "Found on company About page."
}

If no reliable result is found, use:
{
  "first_name": "",
  "last_name": "",
  "title": "",
  "source_url": "",
  "confidence": "",
  "notes": "Reason why name could not be confirmed."
}

Never guess. Never fabricate a name. If uncertain, leave first_name and last_name blank.
"""


def build_search_prompt(company, website):
    domain = domain_of(normalize_url(website)) if website else ""
    domain_search = f"site:{domain} CEO OR founder OR leadership OR \"medical director\"" if domain else ""

    searches = [
        f'"{company}" CEO',
        f'"{company}" founder',
        f'"{company}" "medical director"',
        f'"{company}" LinkedIn CEO OR founder',
        f'"{company}" Crunchbase',
    ]
    if domain_search:
        searches.insert(2, domain_search)

    prompt = f"""Company: {company}
Website: {website or "(none)"}
Domain: {domain or "(none)"}

Please run the following web searches to find the CEO, founder, or medical director:

"""
    for i, s in enumerate(searches, 1):
        prompt += f"{i}. {s}\n"

    prompt += """
After reviewing all search results, extract the most reliable name following the confidence and disambiguation rules in your instructions.

Return ONLY the JSON object described in your instructions.
"""
    return prompt


def research_ceo(client, company, website, retries=3):
    prompt = build_search_prompt(company, website)

    for attempt in range(retries):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=1024,
                system=SYSTEM_PROMPT,
                tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 8}],
                messages=[{"role": "user", "content": prompt}],
            )

            # Extract text from response content blocks
            text = ""
            for block in response.content:
                if hasattr(block, "text"):
                    text += block.text

            # Parse JSON from text
            json_match = re.search(r"\{.*\}", text, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
                return {
                    "first_name":  data.get("first_name", "").strip(),
                    "last_name":   data.get("last_name", "").strip(),
                    "title":       data.get("title", "").strip(),
                    "source_url":  data.get("source_url", "").strip(),
                    "confidence":  data.get("confidence", "").strip(),
                    "notes":       data.get("notes", "").strip(),
                }
            else:
                return {
                    "first_name": "", "last_name": "", "title": "",
                    "source_url": "", "confidence": "",
                    "notes": f"Could not parse model response: {text[:200]}",
                }

        except anthropic.RateLimitError:
            wait = 2 ** (attempt + 2)
            print(f"    Rate limited — waiting {wait}s...")
            time.sleep(wait)
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(4)
            else:
                return {
                    "first_name": "", "last_name": "", "title": "",
                    "source_url": "", "confidence": "",
                    "notes": f"Error: {e}",
                }

    return {
        "first_name": "", "last_name": "", "title": "",
        "source_url": "", "confidence": "",
        "notes": "Exhausted retries.",
    }

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_output(results, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CEO Research"

    headers = [
        "Row #", "Company Name", "Website",
        "CEO First Name", "CEO Last Name", "Title",
        "Source URL", "Confidence", "Notes",
    ]
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r in results:
        conf = r.get("confidence", "")
        row_data = [
            r["row"], r["company"], r["website"],
            r.get("first_name", ""), r.get("last_name", ""), r.get("title", ""),
            r.get("source_url", ""), conf, r.get("notes", ""),
        ]
        ws.append(row_data)
        fill_color = CONFIDENCE_COLORS.get(conf, "FFFFFF")
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        for cell in ws[ws.max_row]:
            cell.fill = fill

    col_widths = [6, 40, 35, 16, 16, 30, 50, 12, 60]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Wrap text in Notes column
    for row in ws.iter_rows(min_row=2, min_col=9, max_col=9):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    wb.save(output_path)
    print(f"\nSaved {len(results)} rows to {output_path}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(input_path, output_path, batch_size, skip_rows):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise SystemExit("ERROR: Set ANTHROPIC_API_KEY environment variable first.")

    client = anthropic.Anthropic(api_key=api_key)

    companies = get_companies(input_path)
    missing = [c for c in companies if not c["existing_first"] or not c["existing_last"]]

    print(f"Total rows: {len(companies)}")
    print(f"Rows missing name: {len(missing)}")

    if skip_rows:
        missing = [c for c in missing if c["row"] not in skip_rows]
        print(f"After skip: {len(missing)} to process")

    batch = missing[:batch_size]
    print(f"Processing batch of {len(batch)}\n")

    results = []
    for i, co in enumerate(batch):
        company = co["company"]
        website = co["website"]
        print(f"[{i+1}/{len(batch)}] Row {co['row']}: {company}")

        # Fast path: name embedded in company field
        first, last, note = extract_name_from_company_field(company)
        if first and last:
            print(f"  → Extracted from company name: {first} {last}")
            results.append({
                **co,
                "first_name":  first,
                "last_name":   last,
                "title":       "Medical Director",
                "source_url":  "",
                "confidence":  "High",
                "notes":       note,
            })
            continue

        # Web research via Claude
        result = research_ceo(client, company, website)
        print(f"  → {result['first_name'] or '—'} {result['last_name'] or '—'} "
              f"[{result['confidence'] or 'no result'}] {result['source_url'] or ''}")
        results.append({**co, **result})
        time.sleep(1.5)  # gentle rate-limit buffer

    write_output(results, output_path)
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Find missing CEO names for clinic spreadsheet.")
    parser.add_argument("--input",   required=True,           help="Input .xlsx file path")
    parser.add_argument("--output",  default="CEO_Results.xlsx", help="Output .xlsx file path")
    parser.add_argument("--batch",   type=int, default=10,    help="Max rows to process (default 10)")
    parser.add_argument("--skip",    type=str, default="",    help="Comma-separated row numbers to skip")
    args = parser.parse_args()

    skip_set = set(int(x.strip()) for x in args.skip.split(",") if x.strip())
    run(args.input, args.output, args.batch, skip_set)
