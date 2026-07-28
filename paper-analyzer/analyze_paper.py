#!/usr/bin/env python3
"""
Extract structured radiation-effects data (TID / SEE / DD) from a PDF research
paper using the OpenAI Responses API.

Standalone replacement for the old server/src/gpt-controller.ts pipeline:
same idea (upload the PDF, ask a sequence of targeted questions, get back
structured JSON), but a single pass, one file, and no threads/polling/website.

Usage:
    cp .env.example .env  # then fill in OPENAI_API_KEY
    python analyze_paper.py paper1.pdf [paper2.pdf ...]
    python analyze_paper.py pdfs/                       # process every PDF in a folder
    python analyze_paper.py pdfs/ --skip-existing        # skip PDFs already extracted
    python analyze_paper.py paper1.pdf --model gpt-4o --outdir results/
"""

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()  # pulls OPENAI_API_KEY (and anything else) from a .env file if present

DEFAULT_MODEL = "gpt-4o"

SYSTEM_INSTRUCTIONS = (
    "You are a radiation-effects specialist for electronic components, well "
    "versed in test methodologies and terminology (TID, SEE, DD). Read the "
    "attached paper carefully and answer with concise, numeric-first data. "
    "If a value truly isn't in the paper, use null rather than guessing."
)

METADATA_PROMPT = """Extract the following details about this research paper:
- title
- year of publication
- all author names
- the paper's main objective, in one or two sentences"""

PARTS_PROMPT = """Identify every distinct electronic part/device tested in this paper.

For each part, report:
- name: part number / device name
- type: e.g. FPGA, memory chip, switching regulator
- manufacturer
- other_details: any other relevant identifying details
- preliminary_test_types: which of TID (Total Ionizing Dose), SEE (Single Event
  Effects: heavy ion / proton / laser / neutron), or DD (Displacement Damage)
  testing was performed on this part. If the same type of test was run as a
  separate campaign more than once, include it once per campaign."""

TEST_PROMPTS = {
    "TID": """Focus only on data related to the part "{part}". Extract Total
Ionizing Dose (TID) test conditions and results:
- maximum fluence, energy levels, facility name, environmental conditions
- whether testing was terrestrial and/or in-flight
- radiation source, maximum TID level reached, dose rate
- whether Enhanced Low Dose Rate Sensitivity (ELDRS) was observed
- dose to failure, whether increased power usage was observed and a
  description of it, how long the failure persisted, and any special notes""",
    "SEE": """Focus only on data related to the part "{part}". Extract Single
Event Effects (SEE) test conditions and results:
- maximum fluence, energy levels, facility name, environmental conditions
- whether testing was terrestrial and/or in-flight
- radiation source, the type of SEE observed (e.g. SEU, SET, SEFI, SEL, SEB,
  SEGR, MBU), amplitude and duration of the effect
- cross-section saturation value, cross-section threshold value, cross-section
  type, and any special notes""",
    "DD": """Focus only on data related to the part "{part}". Extract
Displacement Damage (DD) test conditions and results:
- maximum fluence, energy levels, facility name, environmental conditions
- whether testing was terrestrial and/or in-flight
- radiation source, level of damage observed, a description of the damage
  effects, and any special notes""",
}

METADATA_SCHEMA = {
    "type": "object",
    "properties": {
        "name": {"type": "string"},
        "year": {"type": "integer"},
        "authors": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {"name": {"type": "string"}},
                "required": ["name"],
                "additionalProperties": False,
            },
        },
        "objective": {"type": "string"},
    },
    "required": ["name", "year", "authors", "objective"],
    "additionalProperties": False,
}

PARTS_SCHEMA = {
    "type": "object",
    "properties": {
        "parts": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "type": {"type": "string"},
                    "manufacturer": {"type": "string"},
                    "other_details": {"type": "string"},
                    "preliminary_test_types": {
                        "type": "array",
                        "items": {"type": "string", "enum": ["TID", "SEE", "DD"]},
                    },
                },
                "required": [
                    "name",
                    "type",
                    "manufacturer",
                    "other_details",
                    "preliminary_test_types",
                ],
                "additionalProperties": False,
            },
        }
    },
    "required": ["parts"],
    "additionalProperties": False,
}

_NUM_OR_NULL = {"type": ["number", "null"]}
_STR_OR_NULL = {"type": ["string", "null"]}

_COMMON_CONDITIONS = {
    "max_fluence": _NUM_OR_NULL,
    "energy_levels": {"type": "string"},
    "facility_name": {"type": "string"},
    "environmental_conditions": {"type": "string"},
    "terrestrial": {"type": "boolean"},
    "flight": {"type": "boolean"},
    "special_notes": _STR_OR_NULL,
}

TEST_SCHEMAS = {
    "TID": {
        "type": "object",
        "properties": {
            **_COMMON_CONDITIONS,
            "source": {
                "type": "string",
                "enum": ["Co60", "Protons", "Electrons", "Heavy ions", "X-rays", "Pions"],
            },
            "max_tid": _NUM_OR_NULL,
            "dose_rate": _NUM_OR_NULL,
            "eldrs": {"type": "boolean"},
            "dose_to_failure": _NUM_OR_NULL,
            "increased_power_usage": {"type": "boolean"},
            "power_usage_description": _STR_OR_NULL,
            "failing_time": _STR_OR_NULL,
        },
        "required": [
            *_COMMON_CONDITIONS.keys(),
            "source",
            "max_tid",
            "dose_rate",
            "eldrs",
            "dose_to_failure",
            "increased_power_usage",
            "power_usage_description",
            "failing_time",
        ],
        "additionalProperties": False,
    },
    "SEE": {
        "type": "object",
        "properties": {
            **_COMMON_CONDITIONS,
            "source": {
                "type": "string",
                "enum": ["Heavy ions", "Protons", "Laser", "Neutron", "Electron", "X-rays"],
            },
            "type": {"type": "string"},
            "amplitude": _NUM_OR_NULL,
            "duration": _NUM_OR_NULL,
            "cross_section_saturation": _NUM_OR_NULL,
            "cross_section_threshold": _NUM_OR_NULL,
            "cross_section_type": {"type": "string"},
        },
        "required": [
            *_COMMON_CONDITIONS.keys(),
            "source",
            "type",
            "amplitude",
            "duration",
            "cross_section_saturation",
            "cross_section_threshold",
            "cross_section_type",
        ],
        "additionalProperties": False,
    },
    "DD": {
        "type": "object",
        "properties": {
            **_COMMON_CONDITIONS,
            "source": {"type": "string", "enum": ["Protons", "Neutrons"]},
            "damage_level": _STR_OR_NULL,
            "damage_description": {"type": "string"},
        },
        "required": [
            *_COMMON_CONDITIONS.keys(),
            "source",
            "damage_level",
            "damage_description",
        ],
        "additionalProperties": False,
    },
}


def upload_pdf(client: OpenAI, path: Path) -> str:
    with open(path, "rb") as f:
        file = client.files.create(file=f, purpose="user_data")
    return file.id


def ask(client: OpenAI, model: str, file_id: str, prompt: str, schema_name: str, schema: dict) -> dict:
    response = client.responses.create(
        model=model,
        instructions=SYSTEM_INSTRUCTIONS,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_file", "file_id": file_id},
                    {"type": "input_text", "text": prompt},
                ],
            }
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": schema_name,
                "schema": schema,
                "strict": True,
            }
        },
    )
    return json.loads(response.output_text)


def analyze_pdf(client: OpenAI, path: Path, model: str) -> dict:
    print(f"[{path.name}] uploading...")
    file_id = upload_pdf(client, path)

    print(f"[{path.name}] extracting paper metadata...")
    metadata = ask(client, model, file_id, METADATA_PROMPT, "paper_metadata", METADATA_SCHEMA)

    print(f"[{path.name}] identifying parts and test types...")
    parts = ask(client, model, file_id, PARTS_PROMPT, "parts", PARTS_SCHEMA)["parts"]

    for part in parts:
        part["tids"], part["sees"], part["dds"] = [], [], []
        for test_type in part["preliminary_test_types"]:
            if test_type not in TEST_PROMPTS:
                print(f"[{path.name}]   skipping unrecognized test type '{test_type}' for {part['name']}")
                continue
            print(f"[{path.name}]   extracting {test_type} data for {part['name']}...")
            prompt = TEST_PROMPTS[test_type].format(part=part["name"])
            result = ask(client, model, file_id, prompt, f"{test_type.lower()}_data", TEST_SCHEMAS[test_type])
            {"TID": part["tids"], "SEE": part["sees"], "DD": part["dds"]}[test_type].append(result)

    return {**metadata, "parts": parts}


def collect_pdfs(paths: list[Path]) -> list[Path]:
    pdfs = []
    for path in paths:
        if not path.exists():
            print(f"error: {path} not found", file=sys.stderr)
        elif path.is_dir():
            found = sorted(p for p in path.iterdir() if p.suffix.lower() == ".pdf")
            if not found:
                print(f"warning: no PDFs found in {path}", file=sys.stderr)
            pdfs.extend(found)
        else:
            pdfs.append(path)
    return pdfs


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdfs", nargs="+", type=Path, help="PDF file(s) and/or directories of PDFs to analyze")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"OpenAI model to use (default: {DEFAULT_MODEL})")
    parser.add_argument("--outdir", type=Path, default=None, help="Directory for output JSON (default: next to each PDF)")
    parser.add_argument("--skip-existing", action="store_true", help="Skip PDFs that already have an *_extracted.json output")
    args = parser.parse_args()

    client = OpenAI()  # reads OPENAI_API_KEY from the environment

    for pdf_path in collect_pdfs(args.pdfs):
        outdir = args.outdir or pdf_path.parent
        out_path = outdir / f"{pdf_path.stem}_extracted.json"

        if args.skip_existing and out_path.exists():
            print(f"[{pdf_path.name}] skipping, {out_path} already exists")
            continue

        try:
            data = analyze_pdf(client, pdf_path, args.model)
        except Exception as exc:
            print(f"[{pdf_path.name}] failed: {exc}", file=sys.stderr)
            continue

        outdir.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(data, indent=2))
        print(f"[{pdf_path.name}] done -> {out_path}")


if __name__ == "__main__":
    main()
