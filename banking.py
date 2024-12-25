import os
from typing import Any, Dict
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

        merged_values.append((*formatted_row[0:2], formula_row[2], formatted_row[3], *formula_row[5:-1], formatted_row[-1]))

    return merged_values

def update_sheet_data(client: Client, spreadsheet_id: str, sheet_name: str, ranges: str, transactions: list[Any]):
    sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
    sheet.update(range_name=ranges, values=transactions, value_input_option=ValueInputOption.user_entered)

def clean_cell_value(chars: list[str], value: str) -> str:
    for r in chars:
        value = value.replace(r, "")
    return value

def filter_cell_value(expr: str) -> tuple[list[str], list[str]]:
    cell_in, cell_out = [], []
    for value in expr.split("+"):
        for i, m in enumerate(value.split("-")):
            if i == 0:
                cell_out.append(m)
            else:
                cell_in.append(m)

    return (cell_in, cell_out)

def to_float(s: str) -> float:
    return 0 if s == "" else float(s.replace(",", ""))

def append_category_transaction(transactions: list[Any], date: str, t: str | None, amount: list[str], category: str | None, note: str):
    for a in amount:
        transactions.append([None, date, None, t, f"={a}", category, None, note])

def append_transaction(transactions: list[Any], category_mappers: Dict[str, Dict[str, str | None]], headers: list[str], date: str, micro_transactions: list[Any], note: str):
    # split each incomes into list
    micro_transactions[1] = [] if micro_transactions[1] == "" else [*micro_transactions[1].split("+")]

    # split each expenses into their own list
    for i, categorised in enumerate(micro_transactions[3:]):
        cell_in, cell_out = filter_cell_value(categorised)

        # there are some minuses in the expenses, idk why i did that but it must be put in the income part
        micro_transactions[1] += cell_in 

        micro_transactions[i + 3] = [] if categorised == "" else cell_out

    # append list of incomes into transaction
    _, net_in, _, *categories = micro_transactions
    append_category_transaction(transactions, date, "Income", net_in, None, note)

    # append list of expenses into transaction
    for i, category in enumerate(categories):
        append_category_transaction(transactions, date, category_mappers[headers[i]]["t"], category, category_mappers[headers[i]]["category"], note)

if __name__ == "__main__":
    SPREADSHEETS_FROM = os.environ["SPREADSHEETS_FROM"]
    SHEET_NAME_FROM = "SCB"
    SHEET_RANGE_FROM = "A1:K"

    SPREADSHEETS_TO = os.environ["SPREADSHEETS_TO"]
    SHEET_NAME_TO = "SCB"
    SHEET_RANGE_TO = "A2:H"

    CATEGORY_MAPPER: Dict[str, Dict[str, str | None]] = {
        # KBANK
        "Food": {
            "t": "Expense",
            "category": "Food",
        },
        "Drinks": {
            "t": "Expense",
            "category": "Beverage",
        },
        "Sweets": {
            "t": "Expense",
            "category": "Sweets",
        },
        "Transfer": {
            "t": "Transfer",
            "category": None,
        },
        "Misc": {
            "t": "Expense",
            "category": None,
        },
        # SCB
        "Funding": {
            "t": "Expense",
            "category": "Investment",
        },
        "Subscription": {
            "t": "Expense",
            "category": "Subscription",
        },
        "Debt": {
            "t": "Expense",
            "category": None,
        },
        "Transport": {
            "t": "Expense",
            "category": "Transportation",
        }
    }

    print("Logging in service account...")

    credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.auth.authorize(credentials)

    print("Getting spreadsheets data...")

    values_iter = iter(get_sheet_data(client, SPREADSHEETS_FROM, SHEET_NAME_FROM, SHEET_RANGE_FROM))

    print("Processing data...")

    # headers
    headers = next(values_iter)

    transactions = []

    # initial wallet amount
    date, init, *_ = next(values_iter)
    transactions.append([None, date, f"={to_float(init)}", None, None, None, None, None])

    for row in values_iter:
        date, *micro_transactions, note = row

        micro_transactions = [clean_cell_value(["=", "(", ")"], str(trans)) for trans in micro_transactions]

        _, net_in, net_out, *_ = micro_transactions
        if len(net_in) > 0 or to_float(net_out) > 0:
            append_transaction(transactions, CATEGORY_MAPPER, headers[4:-1], date, micro_transactions, note)

    print("Updating spreadsheets...")

    update_sheet_data(client, SPREADSHEETS_TO, SHEET_NAME_TO, SHEET_RANGE_TO, transactions)

    # for transaction in transactions:
    #     print(transaction)

