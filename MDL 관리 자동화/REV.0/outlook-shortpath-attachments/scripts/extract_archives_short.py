#!/usr/bin/env python3
"""Extract ZIP files to short paths while preserving original names in a manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import zipfile
from pathlib import Path, PurePosixPath


def short_hash(text: str, length: int = 10) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="surrogatepass")).hexdigest()[:length]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("root", type=Path, help="Short attachment output root")
    parser.add_argument("--apply", action="store_true", help="Write extracted files")
    parser.add_argument("--max-files", type=int, default=2000)
    args = parser.parse_args()

    root = args.root.resolve()
    archives = sorted(root.rglob("*.zip"))
    plan: list[dict[str, object]] = []
    extracted = 0
    failed = 0

    for archive_index, archive in enumerate(archives, start=1):
        relative = str(archive.relative_to(root))
        archive_dir = root / "_zip" / f"z{archive_index:03d}_{short_hash(relative, 8)}"
        try:
            with zipfile.ZipFile(archive) as zf:
                members = [item for item in zf.infolist() if not item.is_dir()]
                if len(plan) + len(members) > args.max_files:
                    raise RuntimeError(f"archive file count exceeds --max-files {args.max_files}")
                for member_index, member in enumerate(members, start=1):
                    original = member.filename
                    leaf = PurePosixPath(original.replace("\\", "/")).name
                    extension = Path(leaf).suffix.lower()
                    saved = archive_dir / f"f{member_index:04d}_{short_hash(original, 8)}{extension}"
                    record = {
                        "archive": str(archive),
                        "originalMember": original,
                        "savedPath": str(saved),
                        "uncompressedSize": member.file_size,
                        "result": "planned",
                        "reason": None,
                    }
                    if args.apply:
                        archive_dir.mkdir(parents=True, exist_ok=True)
                        with zf.open(member) as source, saved.open("wb") as destination:
                            shutil.copyfileobj(source, destination)
                        record["result"] = "extracted"
                        extracted += 1
                    plan.append(record)
        except Exception as exc:
            failed += 1
            plan.append({
                "archive": str(archive), "originalMember": None, "savedPath": None,
                "uncompressedSize": None, "result": "failed", "reason": str(exc),
            })

    if args.apply:
        manifest = root / "archive_manifest.json"
        manifest.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({
        "mode": "apply" if args.apply else "dry-run",
        "archives": len(archives),
        "membersFound": sum(item["originalMember"] is not None for item in plan),
        "membersPlanned": sum(item["result"] == "planned" for item in plan),
        "extracted": extracted,
        "failedArchives": failed,
        "root": str(root),
    }, ensure_ascii=False))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
