from __future__ import annotations

import argparse
import re
from pathlib import Path

import requests


def main() -> int:
    parser = argparse.ArgumentParser(description="Download Samolet frontend bundles and search tender API hints")
    parser.add_argument("html_path")
    parser.add_argument("--out", default="scratch_samolet_js")
    args = parser.parse_args()

    html = Path(args.html_path).read_text(encoding="utf-8", errors="ignore")
    output_dir = Path(args.out)
    output_dir.mkdir(parents=True, exist_ok=True)

    bundle_paths = sorted(set(re.findall(r'src="(/[^"]+?\.bundle\.js)"', html)))
    print(f"bundles={len(bundle_paths)}")
    for bundle_path in bundle_paths:
        target = output_dir / Path(bundle_path).name
        if not target.exists():
            response = requests.get(f"https://partner.samolet.ru{bundle_path}", timeout=20)
            response.raise_for_status()
            target.write_bytes(response.content)
        print(f"{target.name}\t{target.stat().st_size}")

    patterns = (
        "fileserializer",
        "preview/file",
        "tenders",
        "attachments",
        "documents",
        "lots",
        "api/",
    )
    for path in sorted(output_dir.glob("*.js")):
        text = path.read_text(encoding="utf-8", errors="ignore")
        hits = [pattern for pattern in patterns if pattern in text]
        if hits:
            print(f"HITS\t{path.name}\t{', '.join(hits)}")
            for match in re.finditer(r"[^\"'`]{0,80}(?:fileserializer|preview/file|api/|tenders|attachments|documents|lots)[^\"'`]{0,120}", text):
                snippet = re.sub(r"\s+", " ", match.group(0)).strip()
                print(f"  {snippet[:260]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
