#!/usr/bin/env python3
"""Lint a carousel Markdown input for mechanical style rules. Zero external dependencies."""

import re
import sys
from pathlib import Path


def parse_front_matter(text):
    if not text.startswith("---\n"):
        return {}, text
    close = text.find("\n---\n", 4)
    if close == -1:
        return {}, text
    meta = {}
    for line in text[4:close].split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        sep = line.find(":")
        if sep == -1:
            continue
        key = line[:sep].strip()
        value = line[sep + 1:].strip().strip("\"'")
        meta[key] = value
    return meta, text[close + 5:]


def strip_directives_content(body, directive):
    """Remove content inside a specific directive block."""
    pattern = re.compile(
        rf"^:::{directive}(?:\s.*)?$\n(.*?\n)^:::$",
        re.MULTILINE | re.DOTALL,
    )
    return pattern.sub("", body)


def strip_code_blocks(body):
    return re.sub(r"^(`{3,}|~{3,}).*?\n.*?^\1\s*$", "", body, flags=re.MULTILINE | re.DOTALL)


def split_sections(body):
    """Split body into page-like sections by :::section or :::pagebreak."""
    parts = re.split(r"^:::(?:section|pagebreak)(?:\s.*)?$", body, flags=re.MULTILINE)
    return [p for p in parts if p.strip()]


def count_numbers(text):
    """Count numeric tokens (integers, decimals, percentages) in text."""
    clean = strip_code_blocks(text)
    clean = re.sub(r"^:::\w+.*$", "", clean, flags=re.MULTILINE)
    clean = re.sub(r"^:::$", "", clean, flags=re.MULTILINE)
    return len(re.findall(r"\d[\d,.]*%?", clean))


def lint(filepath):
    text = Path(filepath).read_text(encoding="utf-8")
    text = text.replace("\r\n", "\n")
    meta, body = parse_front_matter(text)
    findings = []

    # --- Rule: no "你" in body text ---
    body_no_code = strip_code_blocks(body)
    for i, line in enumerate(body_no_code.split("\n"), start=1):
        if "你" in line:
            occurrences = line.count("你")
            findings.append(f'[STYLE] body contains "你" ({occurrences}x): {line.strip()[:80]}')

    # --- Rule: kicker must use neutral words ---
    banned_kicker = ["解读", "深度分析", "研判", "点评"]
    kicker = meta.get("kicker", "")
    for word in banned_kicker:
        if word in kicker:
            findings.append(f'[STYLE] kicker contains subjective word "{word}": {kicker}')

    # --- Rule: no rigid Chinese numbering as main structure ---
    cn_number_lines = []
    for i, line in enumerate(body.split("\n"), start=1):
        if re.match(r"^(?:一|二|三|四|五|六|七|八|九|十)[、，,.]", line.strip()):
            cn_number_lines.append(i)
    if len(cn_number_lines) >= 3:
        findings.append(f"[STYLE] rigid Chinese numbering detected on {len(cn_number_lines)} lines (lines {', '.join(map(str, cn_number_lines[:5]))})")

    # --- Rule: ~3 numbers per section ---
    sections = split_sections(body)
    for idx, section in enumerate(sections, start=1):
        count = count_numbers(section)
        if count > 5:
            preview = section.strip().split("\n")[0][:60]
            findings.append(f"[DENSITY] section {idx} has {count} numeric tokens (recommended ~3): {preview}")

    # --- Rule: no H1 in body (title belongs on cover only) ---
    for i, line in enumerate(body.split("\n"), start=1):
        if re.match(r"^#\s+", line):
            findings.append(f'[STRUCTURE] H1 heading in body (line {i}): {line.strip()[:80]}')

    # --- Rule: risk content must be inside :::risk ---
    risk_keywords = ["风险", "下跌", "暴跌", "亏损", "爆仓", "崩盘", "套牢", "追高"]
    body_outside_risk = strip_directives_content(body_no_code, "risk")
    for keyword in risk_keywords:
        for i, line in enumerate(body_outside_risk.split("\n"), start=1):
            if keyword in line:
                directive_line = line.strip()
                if directive_line.startswith(":::"):
                    continue
                findings.append(f'[COMPLIANCE] "{keyword}" outside :::risk: {directive_line[:80]}')

    # --- Output ---
    if findings:
        print(f"Found {len(findings)} issue(s):\n")
        for f in findings:
            print(f"  {f}")
        print("\nThese are mechanical checks. Review each finding and decide whether to fix.")
    else:
        print("No mechanical issues found.")

    return len(findings)


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] in ("--help", "-h"):
        print("Usage: python lint.py <input.md>")
        sys.exit(0 if sys.argv[1:] else 1)
    count = lint(sys.argv[1])
    sys.exit(0)
