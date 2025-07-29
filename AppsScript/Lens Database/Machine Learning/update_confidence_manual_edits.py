#!/usr/bin/env python3
"""
update_confidence_manual_edits.py
=================================

Re-calculates confidence scores for each row in `Manual Edits.csv` based on
how much of the *original name* string is covered by the user-supplied core
fields.

Core fields considered (mirrors `SimpleLensParser.calculate_confidence_score`):
• Manufacturer
• Series
• Focal Length
• T-Stop
• Prime / Zoom / Special (lens type)
• Notes (optional)

Outputs a new CSV `Manual Edits_confidence_updated.csv` so the original file
remains untouched.
"""

import pandas as pd
import re
from pathlib import Path

# Config
INPUT_FILE = Path(__file__).with_name("Manual Edits.csv")
OUTPUT_FILE = Path(__file__).with_name("Manual Edits_confidence_updated.csv")
CONFIDENCE_FIELD = "Confidence Score"
NEEDS_REVIEW_FIELD = "Needs Review"
CONF_THRESHOLD = 0.6  # below this row is flagged as needing review

# Characters to strip when comparing substrings
CLEAN_REGEX = re.compile(r"[\s\-/()]")


def clean(text: str) -> str:
    """Lowercase and strip spaces / dashes / slashes / parens for fair compare."""
    return CLEAN_REGEX.sub("", str(text).lower())


def calculate_confidence(original: str, manufacturer: str, series: str, focal: str,
                          t_stop: str, lens_type: str, notes: str) -> float:
    """Return fraction of characters from original matched by provided fields."""

    if not original:
        return 0.0

    original_clean = clean(original)
    total_chars = len(original_clean)
    if total_chars == 0:
        return 0.0

    matched = 0
    # account for mm suffix on focal length
    for label, value in zip([
        "manufacturer", "series", "focal", "t_stop", "lens_type", "notes"],
        (manufacturer, series, focal, t_stop, lens_type, notes)):
        if not value or str(value).lower() == "nan":
            continue
        segment = clean(value)
        # Special case focal length
        if label == "focal":
            focal_with_mm = f"{segment}mm"
            if focal_with_mm in original_clean:
                matched += len(focal_with_mm)
                continue
        if segment and segment in original_clean:
            matched += len(segment)

    return min(matched / total_chars, 1.0)


def main():
    if not INPUT_FILE.exists():
        print(f"Input file not found: {INPUT_FILE}")
        return

    df = pd.read_csv(INPUT_FILE)
    if "Original Name" not in df.columns:
        print("Missing 'Original Name' column in Manual Edits.csv")
        return

    # Recalculate confidence per row
    new_scores = []
    needs_review_flags = []
    for _, row in df.iterrows():
        score = calculate_confidence(
            original=str(row.get("Original Name", "")),
            manufacturer=str(row.get("Manufacturer", "")),
            series=str(row.get("Series", "")),
            focal=str(row.get("Focal Length", "")),
            t_stop=str(row.get("T-Stop", "")),
            lens_type=str(row.get("Prime / Zoom / Special", "")),
            notes=str(row.get("Notes", "")),
        )
        new_scores.append(round(score, 6))
        needs_review_flags.append(score < CONF_THRESHOLD)

    df[CONFIDENCE_FIELD] = new_scores
    df[NEEDS_REVIEW_FIELD] = needs_review_flags

    df.to_csv(OUTPUT_FILE, index=False)
    print(f"Updated confidence written to {OUTPUT_FILE} (threshold={CONF_THRESHOLD})")


if __name__ == "__main__":
    main() 