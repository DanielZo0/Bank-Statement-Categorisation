from pdf_statement_processor import parse_wamo_statement_text


def test_parse_wamo_statement_text_parses_signed_amounts():
    sample_text = """Statement
Description Incoming Outgoing Amount
Salary payment
1 September 2025 Payment from client 1,500.00 5,000.00
Card purchase
2 September 2025 Coffee shop -4.50 4,995.50
Closing Balance
"""
    df = parse_wamo_statement_text(sample_text)

    assert len(df) == 2
    assert df.iloc[0]["Amount"] == 1500.0
    assert df.iloc[1]["Amount"] == -4.5
    assert "Salary payment" in df.iloc[0]["Detail"]
