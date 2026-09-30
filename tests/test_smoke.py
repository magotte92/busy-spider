"""Smoke checks for the pinned runtime. These do not crawl the network."""

import py_compile
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_sources_compile():
    py_compile.compile(ROOT / "main.py", doraise=True)
    py_compile.compile(ROOT / "to_excel.py", doraise=True)


def test_main_empty_queue_exits(tmp_path):
    """The committed entry URL is empty, so the crawl loop exits immediately."""
    result = subprocess.run(
        [sys.executable, "main.py"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )
    assert result.returncode == 0
    assert result.stdout.strip() == "0"
    assert not (ROOT / "test.csv").exists()


def test_pandas_ops_used_by_scripts(tmp_path):
    import pandas as pd

    mail_db = pd.DataFrame(columns=["emails", "from"])
    mail_db.loc[0, "emails"] = "a@example.com"
    mail_db.loc[0, "from"] = str('http://example.com/"page').replace('"', "")
    mail_db.dropna(subset=["emails"], inplace=True)
    mail_db.drop_duplicates(subset=["emails"], inplace=True)
    out = tmp_path / "test.csv"
    mail_db.to_csv(out, sep=";", encoding="utf-8", index=False)

    loaded = pd.read_csv(out, sep=";")
    assert list(loaded.columns) == ["emails", "from"]
    assert loaded.loc[0, "emails"] == "a@example.com"
    assert loaded.loc[0, "from"] == "http://example.com/page"


def test_xlsxwriter_ops_used_by_to_excel(tmp_path):
    import xlsxwriter

    path = tmp_path / "hyperlink.xlsx"
    workbook = xlsxwriter.Workbook(str(path))
    worksheet = workbook.add_worksheet("New Speadsheet")
    worksheet.set_column("B:B")
    worksheet.write_url(
        "B1",
        "http://example.com",
        string="Location of the email in google",
    )
    worksheet.write("A1", "a@example.com")
    workbook.close()
    assert path.stat().st_size > 0
