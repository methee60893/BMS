from __future__ import annotations

from pathlib import Path


SRC_DIR = Path(r"D:\CIE\BMS\database")
OUT_DIR = SRC_DIR / "test"
FILES = [
    "01_create_tables.sql",
    "02_create_views.sql",
    "03_create_stored_procedures.sql",
    "04_seed_master_data_template.sql",
    "05_post_deploy_verify.sql",
]


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for file_name in FILES:
        src = SRC_DIR / file_name
        text = src.read_text(encoding="utf-8")
        text = text.replace("USE [BMS];", "USE [BMS_TEST];")
        out = OUT_DIR / file_name.replace(".sql", "_test.sql")
        out.write_text(text, encoding="utf-8")
        print(f"Generated: {out}")


if __name__ == "__main__":
    main()
