#!/usr/bin/env python3
"""
STARR-LAB Radiation Paper Analyzer — Local Model Edition
==========================================================
Same as rad_paper_analyzer.py but runs against a LOCAL LLM via Ollama
(or any OpenAI-compatible local server like vLLM, llama.cpp, LM Studio).

No API key needed. No data leaves your machine.

Requirements:
    pip install openai pymupdf openpyxl

Setup (Ollama — recommended):
    1. Install Ollama:          https://ollama.com/download
    2. Pull a model:            ollama pull qwen2.5:32b
    3. Place PDFs in ./pdfs/
    4. Run:                     python rad_paper_analyzer_local.py

The script talks to Ollama's OpenAI-compatible endpoint at localhost:11434.
For other servers (vLLM, llama.cpp, LM Studio), use --api-base to point
at their endpoint.
"""

import os
import sys
import json
import glob
import re
import time
import argparse
from pathlib import Path

import pymupdf
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ─── Model recommendations for 24 GB VRAM (RTX 3090 / 4090) ───
MODEL_RECOMMENDATIONS = """
┌─────────────────────────────┬───────────┬─────────┬──────────────────────────┐
│ Model (ollama pull ...)     │ VRAM ~    │ Quality │ Notes                    │
├─────────────────────────────┼───────────┼─────────┼──────────────────────────┤
│ qwen2.5:32b                │ ~20 GB    │ ★★★★★  │ Best quality, tight fit   │
│ qwen2.5:14b                │ ~10 GB    │ ★★★★   │ Faster, more headroom     │
│ mistral-small:24b           │ ~14 GB    │ ★★★★   │ Strong all-rounder        │
│ llama3.1:8b                │ ~5 GB     │ ★★★    │ Fast, but less accurate   │
│ qwen3:30b-a3b              │ ~19 GB    │ ★★★★   │ MoE — fast + capable      │
│ gemma3:27b                 │ ~17 GB    │ ★★★★   │ Good at extraction        │
└─────────────────────────────┴───────────┴─────────┴──────────────────────────┘

For best results with radiation paper extraction, qwen2.5:32b or qwen2.5:14b
are recommended. They handle structured JSON output and technical content well.
"""


# ─── Prompts (tuned for local models — shorter, more explicit) ───

SYSTEM_PROMPT = """You are a radiation-effects specialist who extracts data from research papers.
You ALWAYS respond with a single valid JSON object. No other text, no markdown, no explanation."""

EXTRACTION_PROMPT = """Extract radiation test data from this paper into JSON.

Return this exact JSON structure (use null for missing values):

{
  "paper_name": "string",
  "year": 0,
  "authors": ["string"],
  "objective": "string",
  "parts": [
    {
      "name": "part number or device name",
      "type": "component type (FPGA, ADC, memory, etc.)",
      "manufacturer": "string",
      "other_details": "string or null",
      "tests": [
        {
          "test_type": "TID or SEE or DD",
          "source": "radiation source or null",
          "max_fluence": "string or null",
          "energy_levels": "string or null",
          "facility_name": "string or null",
          "environmental_conditions": "string or null",
          "terrestrial": true,
          "flight": false,
          "max_tid": "string or null",
          "dose_rate": "string or null",
          "eldrs": null,
          "dose_to_failure": "string or null",
          "increased_power_usage": null,
          "power_usage_description": "string or null",
          "failing_time": "string or null",
          "see_type": "SEU, SEL, SET, SEFI, etc. or null",
          "amplitude": "string or null",
          "duration": "string or null",
          "cross_section_saturation": "string or null",
          "cross_section_threshold": "string or null",
          "cross_section_type": "string or null",
          "damage_level": "string or null",
          "damage_description": "string or null",
          "special_notes": "string or null"
        }
      ]
    }
  ]
}

Rules:
- One entry per component in "parts". One entry per test in "tests".
- Use null for values not found in the paper.
- Return ONLY the JSON object. No other text.

Paper text:

"""


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from a PDF using PyMuPDF."""
    doc = pymupdf.open(pdf_path)
    text_parts = []
    for page in doc:
        text_parts.append(page.get_text())
    doc.close()
    return "\n".join(text_parts)


def extract_json_from_response(text: str) -> dict | None:
    """Try multiple strategies to pull valid JSON from a model response."""
    if not text:
        return None
    text = text.strip()

    # Strategy 1: direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Strategy 2: strip markdown fences
    cleaned = re.sub(r"^```(?:json)?\s*", "", text)
    cleaned = re.sub(r"\s*```\s*$", "", cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Strategy 3: find the first { ... } block (greedy)
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # Strategy 4: find the outermost braces more carefully
    depth = 0
    start = None
    for i, ch in enumerate(text):
        if ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start is not None:
                try:
                    return json.loads(text[start : i + 1])
                except json.JSONDecodeError:
                    start = None

    return None


def analyze_paper(
    client: OpenAI,
    text: str,
    model: str,
    max_chars: int,
    max_retries: int = 3,
    timeout: float = 300,
) -> dict | None:
    """Send paper text to local model and get structured JSON back with retries."""
    if len(text) > max_chars:
        text = text[:max_chars] + "\n\n[Paper truncated due to context length limits]"

    for attempt in range(1, max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                temperature=0.1,
                max_tokens=4096,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": EXTRACTION_PROMPT + text},
                ],
                timeout=timeout,
            )
            content = response.choices[0].message.content
            data = extract_json_from_response(content)
            if data and "parts" in data:
                return data
            if attempt < max_retries:
                print(f"    ⚠ Attempt {attempt}: got response but JSON was invalid, retrying...")
        except Exception as e:
            err_str = str(e)
            # Fatal errors — no point retrying
            if "not found" in err_str or "404" in err_str:
                print(f"    ✗ Model not found: {model}")
                print(f"      Pull it first: ollama pull {model}")
                sys.exit(1)
            if attempt < max_retries:
                print(f"    ⚠ Attempt {attempt} error: {e}, retrying...")
            else:
                print(f"    ✗ All {max_retries} attempts failed. Last error: {e}")

    return None


def safe(val) -> str:
    """Convert a value to a display string, handling None/bool/lists."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "Yes" if val else "No"
    if isinstance(val, list):
        return ", ".join(str(v) for v in val)
    return str(val)


def build_spreadsheet(results: list[dict], output_path: str):
    """Write extracted data to a formatted Excel workbook."""
    wb = Workbook()

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="2F5496")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")

    def write_sheet(ws, headers, rows):
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        for r_idx, row in enumerate(rows, 2):
            for col, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=col, value=val)
                cell.border = thin_border
                cell.alignment = wrap

    # ── Papers ──
    ws_papers = wb.active
    ws_papers.title = "Papers"
    paper_rows = []
    for i, r in enumerate(results, 1):
        d = r["data"]
        paper_rows.append([
            i, safe(d.get("paper_name")), d.get("year", ""),
            safe(d.get("authors")), safe(d.get("objective")),
            len(d.get("parts", [])), r["filename"],
        ])
    write_sheet(ws_papers,
        ["#", "Paper Name", "Year", "Authors", "Objective", "# Parts", "Source File"],
        paper_rows)
    for col, w in zip("ABCDEFG", [5, 40, 8, 30, 50, 8, 25]):
        ws_papers.column_dimensions[col].width = w

    # ── Components ──
    ws_parts = wb.create_sheet("Components")
    part_rows = []
    for r in results:
        d = r["data"]
        for part in d.get("parts", []):
            part_rows.append([
                safe(d.get("paper_name")), d.get("year", ""),
                safe(part.get("name")), safe(part.get("type")),
                safe(part.get("manufacturer")), safe(part.get("other_details")),
            ])
    write_sheet(ws_parts,
        ["Paper Name", "Year", "Part Name", "Type", "Manufacturer", "Other Details"],
        part_rows)
    for col, w in zip("ABCDEF", [35, 8, 25, 20, 20, 30]):
        ws_parts.column_dimensions[col].width = w

    # ── TID Data ──
    ws_tid = wb.create_sheet("TID Data")
    tid_rows = []
    for r in results:
        d = r["data"]
        for part in d.get("parts", []):
            for test in part.get("tests", []):
                if test.get("test_type", "").upper() != "TID":
                    continue
                tid_rows.append([
                    safe(d.get("paper_name")), d.get("year", ""),
                    safe(part.get("name")), safe(test.get("source")),
                    safe(test.get("max_fluence")), safe(test.get("energy_levels")),
                    safe(test.get("facility_name")), safe(test.get("environmental_conditions")),
                    safe(test.get("terrestrial")), safe(test.get("flight")),
                    safe(test.get("max_tid")), safe(test.get("dose_rate")),
                    safe(test.get("eldrs")), safe(test.get("dose_to_failure")),
                    safe(test.get("increased_power_usage")), safe(test.get("power_usage_description")),
                    safe(test.get("failing_time")), safe(test.get("special_notes")),
                ])
    write_sheet(ws_tid, [
        "Paper", "Year", "Part", "Source", "Max Fluence", "Energy",
        "Facility", "Environment", "Terrestrial", "Flight",
        "Max TID", "Dose Rate", "ELDRS", "Dose to Failure",
        "Increased Power", "Power Description", "Failing Time", "Notes",
    ], tid_rows)

    # ── SEE Data ──
    ws_see = wb.create_sheet("SEE Data")
    see_rows = []
    for r in results:
        d = r["data"]
        for part in d.get("parts", []):
            for test in part.get("tests", []):
                if test.get("test_type", "").upper() != "SEE":
                    continue
                see_rows.append([
                    safe(d.get("paper_name")), d.get("year", ""),
                    safe(part.get("name")), safe(test.get("source")),
                    safe(test.get("see_type")), safe(test.get("max_fluence")),
                    safe(test.get("energy_levels")),
                    safe(test.get("facility_name")), safe(test.get("environmental_conditions")),
                    safe(test.get("terrestrial")), safe(test.get("flight")),
                    safe(test.get("amplitude")), safe(test.get("duration")),
                    safe(test.get("cross_section_saturation")),
                    safe(test.get("cross_section_threshold")),
                    safe(test.get("cross_section_type")),
                    safe(test.get("special_notes")),
                ])
    write_sheet(ws_see, [
        "Paper", "Year", "Part", "Source", "SEE Type", "Max Fluence", "Energy",
        "Facility", "Environment", "Terrestrial", "Flight",
        "Amplitude", "Duration",
        "Cross-Section Sat.", "Cross-Section Thresh.", "Cross-Section Type", "Notes",
    ], see_rows)

    # ── DD Data ──
    ws_dd = wb.create_sheet("DD Data")
    dd_rows = []
    for r in results:
        d = r["data"]
        for part in d.get("parts", []):
            for test in part.get("tests", []):
                if test.get("test_type", "").upper() != "DD":
                    continue
                dd_rows.append([
                    safe(d.get("paper_name")), d.get("year", ""),
                    safe(part.get("name")), safe(test.get("source")),
                    safe(test.get("max_fluence")), safe(test.get("energy_levels")),
                    safe(test.get("facility_name")), safe(test.get("environmental_conditions")),
                    safe(test.get("terrestrial")), safe(test.get("flight")),
                    safe(test.get("damage_level")), safe(test.get("damage_description")),
                    safe(test.get("special_notes")),
                ])
    write_sheet(ws_dd, [
        "Paper", "Year", "Part", "Source", "Max Fluence", "Energy",
        "Facility", "Environment", "Terrestrial", "Flight",
        "Damage Level", "Damage Description", "Notes",
    ], dd_rows)

    # ── All Tests (flat) ──
    ws_all = wb.create_sheet("All Tests (Flat)")
    flat_rows = []
    for r in results:
        d = r["data"]
        for part in d.get("parts", []):
            for test in part.get("tests", []):
                flat_rows.append([
                    safe(d.get("paper_name")), d.get("year", ""),
                    safe(part.get("name")), safe(part.get("manufacturer")),
                    safe(test.get("test_type")), safe(test.get("source")),
                    safe(test.get("facility_name")), safe(test.get("max_fluence")),
                    safe(test.get("energy_levels")), safe(test.get("environmental_conditions")),
                    safe(test.get("terrestrial")), safe(test.get("flight")),
                    safe(test.get("special_notes")),
                ])
    write_sheet(ws_all, [
        "Paper", "Year", "Part", "Manufacturer", "Test Type", "Source",
        "Facility", "Max Fluence", "Energy", "Environment",
        "Terrestrial", "Flight", "Notes",
    ], flat_rows)

    # Set default column widths on data sheets
    for ws in [ws_tid, ws_see, ws_dd, ws_all]:
        for col_letter in [chr(c) for c in range(ord("A"), ord("A") + ws.max_column)]:
            ws.column_dimensions[col_letter].width = 16

    wb.save(output_path)


def check_ollama_running(base_url: str) -> bool:
    """Quick check if the server is reachable."""
    import urllib.request
    import urllib.error
    try:
        # Strip /v1 for the health check
        health_url = base_url.replace("/v1", "")
        req = urllib.request.Request(health_url, method="GET")
        urllib.request.urlopen(req, timeout=5)
        return True
    except Exception:
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Extract radiation effects data from papers using a LOCAL LLM.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=MODEL_RECOMMENDATIONS,
    )
    parser.add_argument(
        "--pdf-dir", default="pdfs",
        help="Directory containing PDF papers (default: pdfs)",
    )
    parser.add_argument(
        "--output", default="rad_effects_database.xlsx",
        help="Output Excel file (default: rad_effects_database.xlsx)",
    )
    parser.add_argument(
        "--model", default="qwen2.5:32b",
        help="Model name as known by your server (default: qwen2.5:32b)",
    )
    parser.add_argument(
        "--api-base", default="http://localhost:11434/v1",
        help="OpenAI-compatible API base URL (default: http://localhost:11434/v1 for Ollama)",
    )
    parser.add_argument(
        "--max-chars", type=int, default=30000,
        help="Max characters of paper text to send (default: 30000, ~8k tokens)",
    )
    parser.add_argument(
        "--retries", type=int, default=3,
        help="Retries per paper on JSON parse failure (default: 3)",
    )
    parser.add_argument(
        "--timeout", type=float, default=600,
        help="Timeout in seconds per API call (default: 600 — local models are slow)",
    )
    parser.add_argument(
        "--dump-json", action="store_true",
        help="Also save raw JSON responses per paper",
    )
    parser.add_argument(
        "--progress-dir", default=".progress",
        help="Directory for saving progress (default: .progress). Allows resuming interrupted runs.",
    )
    parser.add_argument(
        "--no-resume", action="store_true",
        help="Start fresh, ignoring any saved progress",
    )
    parser.add_argument(
        "--show-models", action="store_true",
        help="Show recommended models for 24 GB VRAM and exit",
    )
    args = parser.parse_args()

    if args.show_models:
        print(MODEL_RECOMMENDATIONS)
        sys.exit(0)

    # Check server
    if not check_ollama_running(args.api_base):
        print(f"Error: Cannot reach server at {args.api_base}")
        if "11434" in args.api_base:
            print()
            print("  Ollama doesn't appear to be running. Start it with:")
            print("    ollama serve")
            print()
            print(f"  Then pull a model:")
            print(f"    ollama pull {args.model}")
        else:
            print(f"  Make sure your local LLM server is running at {args.api_base}")
        sys.exit(1)

    # Find PDFs
    pdf_dir = Path(args.pdf_dir)
    if not pdf_dir.is_dir():
        print(f"Error: PDF directory '{pdf_dir}' not found.")
        print(f"  mkdir {pdf_dir} && cp /path/to/papers/*.pdf {pdf_dir}/")
        sys.exit(1)

    pdf_files = sorted(glob.glob(str(pdf_dir / "*.pdf")))
    if not pdf_files:
        print(f"Error: No PDF files found in '{pdf_dir}/'")
        sys.exit(1)

    print(f"STARR-LAB Radiation Paper Analyzer (Local)")
    print(f"{'=' * 45}")
    print(f"  Server:    {args.api_base}")
    print(f"  Model:     {args.model}")
    print(f"  PDFs:      {len(pdf_files)} papers in {pdf_dir}/")
    print(f"  Max chars: {args.max_chars:,} (~{args.max_chars // 4:,} tokens)")
    print(f"  Retries:   {args.retries}")
    print(f"  Output:    {args.output}")
    print()

    # Set up progress directory for resume support
    progress_dir = Path(args.progress_dir)
    if args.no_resume and progress_dir.exists():
        import shutil
        shutil.rmtree(progress_dir)
        print("  Cleared previous progress.\n")
    progress_dir.mkdir(exist_ok=True)

    # Load any previously completed results
    results = []
    completed = set()
    if not args.no_resume:
        for pf in sorted(progress_dir.glob("*.json")):
            try:
                with open(pf) as f:
                    saved = json.load(f)
                results.append(saved)
                completed.add(saved["filename"])
            except (json.JSONDecodeError, KeyError):
                pass
        if completed:
            print(f"  Resuming: {len(completed)} paper(s) already done, "
                  f"{len(pdf_files) - len(completed)} remaining\n")

    client = OpenAI(base_url=args.api_base, api_key="not-needed")

    for i, pdf_path in enumerate(pdf_files, 1):
        filename = os.path.basename(pdf_path)

        if filename in completed:
            continue

        print(f"[{i}/{len(pdf_files)}] Processing: {filename}")

        # Extract text
        print(f"    Extracting text from PDF...")
        try:
            text = extract_text_from_pdf(pdf_path)
        except Exception as e:
            print(f"    ✗ Failed to read PDF: {e}")
            continue

        if len(text.strip()) < 100:
            print(f"    ✗ PDF is scanned/image-only (very little text extracted)")
            continue

        print(f"    Extracted {len(text):,} chars ({len(text) // 4:,} est. tokens)")
        if len(text) > args.max_chars:
            print(f"    Truncating to {args.max_chars:,} chars to fit context window")

        # Analyze
        print(f"    Sending to {args.model}...")
        t0 = time.time()
        data = analyze_paper(
            client, text, args.model,
            max_chars=args.max_chars,
            max_retries=args.retries,
            timeout=args.timeout,
        )
        elapsed = time.time() - t0

        if data is None:
            print(f"    ✗ Analysis failed after {args.retries} attempts ({elapsed:.0f}s)")
            continue

        parts_count = len(data.get("parts", []))
        tests_count = sum(len(p.get("tests", [])) for p in data.get("parts", []))
        print(f"    ✓ Done in {elapsed:.1f}s — {parts_count} component(s), {tests_count} test(s)")

        entry = {"filename": filename, "data": data}
        results.append(entry)

        # Save progress immediately so we can resume
        safe_name = filename.replace(".pdf", ".json")
        with open(progress_dir / safe_name, "w") as f:
            json.dump(entry, f, indent=2)

        if args.dump_json:
            json_path = pdf_path.replace(".pdf", "_extracted.json")
            with open(json_path, "w") as f:
                json.dump(data, f, indent=2)

    print()
    if not results:
        print("No papers were successfully processed.")
        sys.exit(1)

    print(f"Building spreadsheet: {args.output}")
    build_spreadsheet(results, args.output)

    total_parts = sum(len(r["data"].get("parts", [])) for r in results)
    total_tests = sum(
        len(t.get("tests", []))
        for r in results
        for t in r["data"].get("parts", [])
    )
    print(f"  ✓ {len(results)} paper(s), {total_parts} component(s), {total_tests} test(s)")
    print(f"  ✓ Saved to {args.output}")
    print()
    print("Done!")


if __name__ == "__main__":
    main()
