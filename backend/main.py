import os
import sys
import argparse
from credit_report import extract_credit_report_data

SUPPORTED_EXTS = {".pdf", ".htm", ".html", ".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp"}


def process_file(path: str, excel_path: str) -> None:
    if not os.path.isfile(path):
        print(f"✗ Not a file: {path}")
        return
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        print(f"✗ Unsupported file type: {path}")
        return
    with open(path, "rb") as f:
        content = f.read()
    try:
        row = extract_credit_report_data(content, os.path.basename(path), excel_path)
        name = row.get("full_name") or "(no name)"
        score = row.get("credit_score") or "N/A"
        print(f"✓ {path} -> {name} | Score: {score}")
    except Exception as e:
        print(f"✗ Failed {path}: {e}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract data from credit reports (PDF or HTML) and append to Excel."
    )
    parser.add_argument("files", nargs="+", help="One or more credit report files")
    parser.add_argument("-o", "--output", default="credit_reports.xlsx",
                        help="Path to Excel output file (default: credit_reports.xlsx)")
    args = parser.parse_args()

    print(f"Output file: {args.output}\n")
    for path in args.files:
        process_file(path, args.output)

    print(f"\nDone. Data saved to {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())