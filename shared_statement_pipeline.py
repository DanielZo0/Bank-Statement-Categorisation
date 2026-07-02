"""
Shared transaction enrichment and Excel export pipeline.
"""

from typing import Tuple

import pandas as pd

from common_categorization import (
    MONTH_COLORS,
    capitalize_first,
    extract_counterparty,
    extract_invoice,
    get_transaction_type,
    limit_length,
)


ACCOUNTING_OUTPUT_COLUMNS = [
    "Type",
    "Account Reference",
    "Nominal A/C Ref",
    "Department Code",
    "Date",
    "reference",
    "Details",
    "Net Amount",
    "Tax Code",
    "Tax Amount",
    "Exchange Rate",
    "Extra Reference",
    "User Name",
    "Project Refn",
    "Cost Code Refn",
    "Invoice",
    "Counterparty",
]


def process_transactions(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Split transactions and add derived/accounting columns."""
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    incoming = df[df["Amount"] >= 0].copy()
    outgoing = df[df["Amount"] < 0].copy()

    for transactions_df in [incoming, outgoing]:
        if transactions_df.empty:
            continue

        transactions_df["Type"] = transactions_df["Detail"].apply(
            lambda x: limit_length(capitalize_first(get_transaction_type(str(x).lower())))
        )
        transactions_df["Invoice"] = transactions_df["Detail"].apply(
            lambda x: limit_length(capitalize_first(extract_invoice(str(x))))
        )
        transactions_df["Counterparty"] = transactions_df["Detail"].apply(
            lambda x: limit_length(capitalize_first(extract_counterparty(str(x))))
        )
        transactions_df["Counterparty"] = transactions_df["Counterparty"].apply(
            lambda x: int(x) if str(x).isdigit() else x
        )

        transactions_df["Account Reference"] = ""
        transactions_df["Nominal A/C Ref"] = ""
        transactions_df["Department Code"] = ""
        transactions_df["reference"] = ""
        transactions_df["Details"] = transactions_df["Detail"]
        transactions_df["Net Amount"] = transactions_df["Amount"].abs()
        transactions_df["Tax Code"] = "T9"
        transactions_df["Tax Amount"] = 0.00
        transactions_df["Exchange Rate"] = ""
        transactions_df["Extra Reference"] = ""
        transactions_df["User Name"] = ""
        transactions_df["Project Refn"] = ""
        transactions_df["Cost Code Refn"] = ""

    incoming = incoming.sort_values("Date").reset_index(drop=True)
    outgoing = outgoing.sort_values("Date").reset_index(drop=True)
    return incoming, outgoing


def _write_accounting_sheet(
    writer: pd.ExcelWriter,
    workbook,
    sheet_name: str,
    table_name: str,
    table_style: str,
    df: pd.DataFrame,
):
    if df.empty:
        pd.DataFrame(columns=ACCOUNTING_OUTPUT_COLUMNS).to_excel(
            writer, sheet_name=sheet_name, index=False
        )
        return

    output = df[ACCOUNTING_OUTPUT_COLUMNS].copy()
    output.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
    worksheet = writer.sheets[sheet_name]

    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    currency_format = workbook.add_format({"num_format": "#,##0.00"})

    worksheet.add_table(0, 0, len(output), len(ACCOUNTING_OUTPUT_COLUMNS) - 1, {
        "name": table_name,
        "style": table_style,
        "columns": [
            {"header": "Type"},
            {"header": "Account Reference"},
            {"header": "Nominal A/C Ref"},
            {"header": "Department Code"},
            {"header": "Date", "format": date_format},
            {"header": "reference"},
            {"header": "Details"},
            {"header": "Net Amount", "format": currency_format},
            {"header": "Tax Code"},
            {"header": "Tax Amount", "format": currency_format},
            {"header": "Exchange Rate"},
            {"header": "Extra Reference"},
            {"header": "User Name"},
            {"header": "Project Refn"},
            {"header": "Cost Code Refn"},
            {"header": "Invoice"},
            {"header": "Counterparty"},
        ],
    })

    for idx, row in output.iterrows():
        if pd.notna(row["Date"]):
            color = MONTH_COLORS.get(row["Date"].month, "#FFFFFF")
            date_cell_format = workbook.add_format(
                {"bg_color": color, "num_format": "yyyy-mm-dd"}
            )
            text_cell_format = workbook.add_format({"bg_color": color})
            currency_cell_format = workbook.add_format(
                {"bg_color": color, "num_format": "#,##0.00"}
            )
            excel_row = idx + 1
            for col_idx, col_name in enumerate(ACCOUNTING_OUTPUT_COLUMNS):
                if col_name == "Date":
                    worksheet.write(excel_row, col_idx, row[col_name], date_cell_format)
                elif col_name in ["Net Amount", "Tax Amount"]:
                    worksheet.write(excel_row, col_idx, row[col_name], currency_cell_format)
                else:
                    worksheet.write(excel_row, col_idx, row[col_name], text_cell_format)

    worksheet.set_column("A:A", 20)
    worksheet.set_column("B:B", 18)
    worksheet.set_column("C:C", 18)
    worksheet.set_column("D:D", 18)
    worksheet.set_column("E:E", 12)
    worksheet.set_column("F:F", 15)
    worksheet.set_column("G:G", 40)
    worksheet.set_column("H:H", 15)
    worksheet.set_column("I:I", 12)
    worksheet.set_column("J:J", 12)
    worksheet.set_column("K:K", 15)
    worksheet.set_column("L:L", 18)
    worksheet.set_column("M:M", 15)
    worksheet.set_column("N:N", 15)
    worksheet.set_column("O:O", 15)
    worksheet.set_column("P:P", 20)
    worksheet.set_column("Q:Q", 26)


def export_to_excel(
    source_df: pd.DataFrame, incoming_df: pd.DataFrame, outgoing_df: pd.DataFrame, output_path: str
):
    """Export source/incoming/outgoing sheets with shared formatting."""
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        workbook = writer.book
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        currency_format = workbook.add_format({"num_format": "#,##0.00"})

        source_df.to_excel(writer, sheet_name="SOURCE", index=False, startrow=0)
        worksheet_source = writer.sheets["SOURCE"]
        if not source_df.empty:
            worksheet_source.add_table(0, 0, len(source_df), 2, {
                "name": "SOURCE_TABLE",
                "style": "Table Style Medium 2",
                "columns": [
                    {"header": "Date", "format": date_format},
                    {"header": "Detail"},
                    {"header": "Amount", "format": currency_format},
                ],
            })
            worksheet_source.set_column("A:A", 12, date_format)
            worksheet_source.set_column("B:B", 70)
            worksheet_source.set_column("C:C", 15, currency_format)

        _write_accounting_sheet(
            writer=writer,
            workbook=workbook,
            sheet_name="INCOMING",
            table_name="INCOMING_TABLE",
            table_style="Table Style Medium 9",
            df=incoming_df,
        )
        _write_accounting_sheet(
            writer=writer,
            workbook=workbook,
            sheet_name="OUTGOING",
            table_name="OUTGOING_TABLE",
            table_style="Table Style Medium 4",
            df=outgoing_df,
        )

    print(f"Excel file created: {output_path}")
