import pandas as pd

from shared_statement_pipeline import export_to_excel, process_transactions


def test_process_transactions_adds_accounting_columns():
    source_df = pd.DataFrame(
        [
            {"Date": pd.Timestamp("2025-09-01"), "Detail": "Salary payment", "Amount": 1500.0},
            {"Date": pd.Timestamp("2025-09-02"), "Detail": "Card transaction coffee", "Amount": -4.5},
        ]
    )

    incoming, outgoing = process_transactions(source_df)

    assert len(incoming) == 1
    assert len(outgoing) == 1
    assert "Net Amount" in incoming.columns
    assert outgoing.iloc[0]["Net Amount"] == 4.5


def test_export_to_excel_creates_file(tmp_path):
    source_df = pd.DataFrame(
        [
            {"Date": pd.Timestamp("2025-09-01"), "Detail": "Salary payment", "Amount": 1500.0},
            {"Date": pd.Timestamp("2025-09-02"), "Detail": "Card transaction coffee", "Amount": -4.5},
        ]
    )
    incoming, outgoing = process_transactions(source_df)
    output_path = tmp_path / "categorized.xlsx"

    export_to_excel(source_df, incoming, outgoing, str(output_path))

    assert output_path.exists()
    assert output_path.stat().st_size > 0
