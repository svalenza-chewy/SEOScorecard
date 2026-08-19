from __future__ import annotations

import csv
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SITE_DIR = ROOT / "site"
WEB_DIR = ROOT / "web"
OUTPUTS_DIR = ROOT / "outputs"
DETAIL_ROW_LIMIT = 500


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True)


def copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def copy_text_with_lf(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    content = source.read_text(encoding="utf-8-sig")
    destination.write_text(content.replace("\r\n", "\n"), encoding="utf-8")


def write_site_index() -> None:
    (SITE_DIR / "index.html").write_text(
        """<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Chewy SEO Scorecard</title>
    <meta http-equiv="refresh" content="0; url=./web/" />
    <style>
      body {
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        color: #17233a;
        background: #f6f8fb;
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      a {
        color: #1f5ea8;
        font-weight: 800;
      }
    </style>
  </head>
  <body>
    <main>
      <h1>Opening Chewy SEO Scorecard</h1>
      <p><a href="./web/">Open the dashboard</a></p>
    </main>
    <script>
      window.location.replace("./web/" + window.location.search + window.location.hash);
    </script>
  </body>
</html>
""",
        encoding="utf-8",
    )
    (SITE_DIR / ".nojekyll").write_text("", encoding="utf-8")


def trim_detail_file(source: Path, destination: Path) -> bool:
    if not source.exists():
        return False

    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.open("r", encoding="utf-8-sig", newline="") as src:
        reader = csv.reader(src)
        with destination.open("w", encoding="utf-8", newline="") as dst:
            writer = csv.writer(dst, lineterminator="\n")
            for index, row in enumerate(reader):
                if index > DETAIL_ROW_LIMIT:
                    break
                writer.writerow(row)
    return True


def copy_issue_summary() -> None:
    source = OUTPUTS_DIR / "page_type_issue_summary.csv"
    destination = SITE_DIR / "outputs" / "page_type_issue_summary.csv"
    detail_destination = SITE_DIR / "outputs" / "issue_details"
    detail_destination.mkdir(parents=True, exist_ok=True)

    with source.open("r", encoding="utf-8-sig", newline="") as src:
        reader = csv.DictReader(src)
        fieldnames = reader.fieldnames or []
        rows = []

        for row in reader:
            detail_file = row.get("detail_file", "").strip()
            if detail_file:
                detail_name = Path(detail_file).name
                copied = trim_detail_file(
                    OUTPUTS_DIR / "issue_details" / detail_name,
                    detail_destination / detail_name,
                )
                row["detail_file"] = f"../outputs/issue_details/{detail_name}" if copied else ""
            rows.append(row)

    with destination.open("w", encoding="utf-8", newline="") as dst:
        writer = csv.DictWriter(dst, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    reset_dir(SITE_DIR)
    shutil.copytree(WEB_DIR, SITE_DIR / "web")
    write_site_index()

    copy_text_with_lf(
        OUTPUTS_DIR / "page_type_heatmap_with_ranking_keywords.csv",
        SITE_DIR / "outputs" / "page_type_heatmap_with_ranking_keywords.csv",
    )
    copy_text_with_lf(
        OUTPUTS_DIR / "url_keyword_summary.csv",
        SITE_DIR / "outputs" / "url_keyword_summary.csv",
    )
    copy_issue_summary()

    print(f"Built {SITE_DIR}")


if __name__ == "__main__":
    main()
