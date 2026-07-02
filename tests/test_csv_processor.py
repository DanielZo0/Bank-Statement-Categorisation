from csv_statement_processor import extract_transactions_from_csv


def test_extract_transactions_from_csv_reads_bov_layout(tmp_path):
    csv_content = """Account Summary
Transaction History
Date,Detail,Amount
01/10/2025,Salary payment,1000.00
02/10/2025,Coffee,-3.50
"""
    csv_file = tmp_path / "statement.csv"
    csv_file.write_text(csv_content, encoding="utf-8")

    df = extract_transactions_from_csv(str(csv_file))

    assert len(df) == 2
    assert list(df.columns) == ["Date", "Detail", "Amount"]
    assert df.iloc[0]["Amount"] == 1000.0
    assert df.iloc[1]["Amount"] == -3.5
