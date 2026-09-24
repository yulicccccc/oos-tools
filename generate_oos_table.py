"""
OOS Table Generator - From Extracted Data to Word Table
=========================================================
Usage: Called by the AI agent after extracting data from scanned PDF images.
       The AI provides a structured data dict, this script fills the template.

Supports: EM, ScanRDI, Celsis, USP71

Architecture:
  - Templates: existing 'tables for em/scan/celsis/71.docx' (source of truth)
  - This script: populates {{ variables }} via docxtpl
  - AI agent: extracts data from scanned PDF images, builds the data dict
"""

import sys
import io
import os
import json
from datetime import datetime, timedelta
from docxtpl import DocxTemplate

# Force UTF-8 output (only when running directly, not when imported)
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Resolve paths relative to this script's directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_MAP = {
    'em':     os.path.join(SCRIPT_DIR, 'tables for em.docx'),
    'scan':   os.path.join(SCRIPT_DIR, 'tables for scan.docx'),
    'celsis': os.path.join(SCRIPT_DIR, 'tables for celsis.docx'),
    'usp71':  os.path.join(SCRIPT_DIR, 'tables for 71.docx'),
}

# Date format used in templates
DATE_FMT = '%d%b%Y'  # e.g., 19JUN2026


def business_day_before(dt):
    """Get the previous business day (skip weekends). Celsis/Scan/USP71 = Mon-Fri only."""
    prev = dt - timedelta(days=1)
    while prev.weekday() >= 5:  # 5=Sat, 6=Sun
        prev -= timedelta(days=1)
    return prev


def business_day_after(dt):
    """Get the next business day (skip weekends). Celsis/Scan/USP71 = Mon-Fri only."""
    nxt = dt + timedelta(days=1)
    while nxt.weekday() >= 5:
        nxt += timedelta(days=1)
    return nxt


def format_date(dt):
    """Format a datetime to DDMMMYYYY uppercase (e.g., 19JUN2026)."""
    return dt.strftime(DATE_FMT).upper()


def generate_table(oos_type: str, data: dict, output_path: str) -> str:
    """
    Generate an OOS table Word document from structured data.

    Args:
        oos_type: One of 'em', 'scan', 'celsis', 'usp71'
        data: Dict of template variables (keys must match {{ var }} in template)
        output_path: Full path for the output .docx file

    Returns:
        The output file path
    """
    oos_type = oos_type.lower().strip()
    if oos_type not in TEMPLATE_MAP:
        raise ValueError(f"Unknown OOS type: {oos_type}. Must be one of: {list(TEMPLATE_MAP.keys())}")

    template_path = TEMPLATE_MAP[oos_type]
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    # Load template
    tpl = DocxTemplate(template_path)

    # Check which variables the template expects
    expected_vars = tpl.get_undeclared_template_variables()
    provided_vars = set(data.keys())

    missing = expected_vars - provided_vars
    if missing:
        print(f"WARNING: Missing variables (will be blank): {sorted(missing)}")

    extra = provided_vars - expected_vars
    if extra:
        print(f"INFO: Extra variables (ignored): {sorted(extra)}")

    # Fill all missing vars with empty string to avoid rendering errors
    context = {var: '' for var in expected_vars}
    context.update(data)

    # Render and save
    tpl.render(context)
    tpl.save(output_path)

    print(f"Generated: {output_path}")
    return output_path


def build_output_filename(oos_type: str, oos_number: str, test_date: str, output_dir: str = None) -> str:
    """
    Build a standard output filename.
    Example: "EM table OOS-261185 08MAY2026.docx"
    """
    type_labels = {
        'em': 'EM',
        'scan': 'Scan',
        'celsis': 'Celsis',
        'usp71': 'USP71',
    }
    label = type_labels.get(oos_type.lower(), oos_type.upper())
    filename = f"{label} table {oos_number} {test_date}.docx"

    if output_dir:
        return os.path.join(output_dir, filename)
    return os.path.join(SCRIPT_DIR, filename)


# ============================================================
# Quick CLI for testing
# ============================================================
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python generate_oos_table.py <oos_type> <json_data_file> [output_path]")
        print("  oos_type: em | scan | celsis | usp71")
        print("  json_data_file: path to JSON file with template variables")
        sys.exit(1)

    oos_type = sys.argv[1]
    json_path = sys.argv[2]

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if len(sys.argv) > 3:
        output = sys.argv[3]
    else:
        output = build_output_filename(
            oos_type,
            data.get('_oos_number', 'UNKNOWN'),
            data.get('_test_date', datetime.now().strftime('%d%b%Y')),
        )

    generate_table(oos_type, data, output)
