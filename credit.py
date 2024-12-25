import os
from typing import Any
import gspread
from google.oauth2.service_account import Credentials
from gspread.client import Client
from gspread.utils import ValueInputOption, ValueRenderOption

SERVICE_ACCOUNT_FILE = "cerds.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_sheet_data(client: Client, spreadsheet_id: str, sheet_name: str, ranges: str):
    sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)

    formatted_values = sheet.get_values(ranges, value_render_option=ValueRenderOption.formatted)
    formula_values = sheet.get_values(ranges, value_render_option=ValueRenderOption.formula)

    merged_values = []
    for i in range(len(formatted_values)):
        formatted_row = formatted_values[i]
        formula_row = formula_values[i]

        merged_values.append((*formatted_row[0:3], formula_row[4], formatted_row[5]))

    return merged_values

def update_sheet_data(client: Client, spreadsheet_id: str, sheet_name: str, ranges: str, transactions: list[Any]):
    sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
    sheet.update(range_name=ranges, values=transactions, value_input_option=ValueInputOption.user_entered)

def clean_cell_value(chars: list[str], value: str) -> str:
    for r in chars:
        value = value.replace(r, "")
    return value

def to_float(s: str) -> float:
    return 0 if s == "" else float(s.replace(",", ""))

def append_transaction(transactions: list[Any], date: str, micro_transactions: list[Any], note: str):
    micro_transactions[2] = micro_transactions[2].split("+")

    splitted_note = [note.strip() for note in note.replace("+", ",").split(",")]
    note_list = splitted_note if len(splitted_note) == len(micro_transactions[2]) else [note] * len(micro_transactions[2])

    for i, transaction in enumerate(micro_transactions[2]):
        transaction = to_float(transaction)
        t = "Pay" if transaction > 0 else "Void"
        transactions.append([None, date, None, t, f"={abs(transaction)}", None, None, note_list[i]])

if __name__ == "__main__":
    SPREADSHEETS_FROM = os.environ["SPREADSHEETS_FROM"]
    SHEET_NAME_FROM = "CREDIT"
    SHEET_RANGE_FROM = "A1:K"

    SPREADSHEETS_TO = os.environ["SPREADSHEETS_TO"]
    SHEET_NAME_TO = "CREDIT"
    SHEET_RANGE_TO = "A2:H"

    print("Logging in service account...")

    credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.auth.authorize(credentials)

    print("Getting spreadsheets data...")

    values_iter = iter(get_sheet_data(client, SPREADSHEETS_FROM, SHEET_NAME_FROM, SHEET_RANGE_FROM))

    print("Processing data...")

    # headers
    _ = next(values_iter)

    transactions = []

    # initial wallet amount
    date, init, *_ = next(values_iter)
    transactions.append([None, date, f"={to_float(init)}", None, None, None, None, None])

    prev_total = 0

    for row in values_iter:
        date, *micro_transactions, note = row
        micro_transactions = [clean_cell_value(["=", "(", ")"], str(trans)) for trans in micro_transactions]

        total, net_out, *_ = micro_transactions
        total = to_float(total)
        net_out = to_float(net_out)

        # credit payout
        if total < prev_total:
            diff = prev_total - total + net_out
            transactions.append([None, date, None, "Transfer", f"={diff}", None, None, ""])

        if net_out != 0:
            append_transaction(transactions, date, micro_transactions, note)

        prev_total = total

    print("Updating spreadsheets...")

    update_sheet_data(client, SPREADSHEETS_TO, SHEET_NAME_TO, SHEET_RANGE_TO, transactions)

    # for transaction in transactions:
    #     print(transaction)

