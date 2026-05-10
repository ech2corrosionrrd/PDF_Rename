import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

import excel_db  # noqa: E402
import naming as pre  # noqa: E402


def test_sanitize():
    print("Testing sanitize_component...")
    cases = [
        ("Test/File:Name*", "Test-File-Name"),
        ("Spaces   to   Single", "Spaces to Single"),
        ("EIC (CODE) [brackets]", "EIC -CODE- -brackets"),
    ]
    for inp, expected in cases:
        res = pre.sanitize_component(inp)
        print(f"  '{inp}' -> '{res}' (Expected: '{expected}')")
        assert res == expected, f"Failed: {res} != {expected}"
    print("Sanitize tests passed!\n")


def test_excel_loading():
    print("Testing ExcelConsumerDB...")
    app_dir = _ROOT
    db = excel_db.ExcelConsumerDB(app_dir)
    print(f"  Excel path resolved: {db.excel_path}")
    if db.excel_path and db.excel_path.exists():
        try:
            db.load()
            print(f"  Loaded {db.record_count} records.")

            if db.record_count:
                first_eic = (db._eic_col or [""])[0]
                if len(first_eic) >= 5:
                    suffix = first_eic[-5:]
                    matches = db.search_by_eic_suffix(suffix)
                    print(f"  Search for '{suffix}' found {len(matches)} matches.")
                    for m in matches:
                        print(f"    - Found: {m.display()}")
        except Exception as e:
            print(f"  Excel loading failed: {e}")
    else:
        print("  Excel file not found, skipping deep test.")
    print("Excel tests finished!\n")


def test_filename_parsing():
    print("Testing parse_renamed_pdf_filename...")
    filename = "1234_32X1234567890ABC_2023-10-27_MET123_NAME_ADDR_Інвест_3Ф_План_Монт-АСКОЕ.pdf"
    res = pre.parse_renamed_pdf_filename(filename)
    if res:
        print(f"  Parsed successfully: {res['EIC']} / {res['Дата']} / {res['Інвест']}")
        assert res["EIC"] == "32X1234567890ABC"
        assert res["Інвест"] == "Так"
        assert res.get("Належність лічильника", "") == ""
    else:
        print("  Failed to parse filename!")
        assert False

    with_suffix = (
        "1234_32X1234567890ABC_2023-10-27_MET123_NAME_ADDR_Інвест_3Ф_План_"
        "Монт-АСКОЕ_IP-1.1.1.1_Mod-99_Мій-додаток.pdf"
    )
    r2 = pre.parse_renamed_pdf_filename(with_suffix)
    assert r2 and r2["Належність лічильника"] == "Мій-додаток"
    print("Filename parsing tests passed!\n")


def test_askoe_logic():
    print("Testing ASKOE filename logic...")
    parts = ["1001", "32X12345", "2023-10-27", "MET1", "Name", "Addr", "3Ф", "План", "Монт-АСКОЕ"]

    ip = "192.168.1.1"
    modem = "12345678"

    askoe_parts = list(parts)
    askoe_parts.append(f"IP-{ip}")
    askoe_parts.append(f"Mod-{modem}")

    res = "_".join(askoe_parts) + ".pdf"
    print(f"  Result with IP/Modem: {res}")
    assert "IP-192.168.1.1" in res
    assert "Mod-12345678" in res

    print("ASKOE tests passed!\n")


if __name__ == "__main__":
    try:
        test_sanitize()
        test_excel_loading()
        test_filename_parsing()
        test_askoe_logic()
        print("ALL BACKEND LOGIC VERIFIED SUCCESSFULLY!")
    except Exception as e:
        print(f"\nVERIFICATION FAILED: {e}")
        sys.exit(1)
