"""
mapping_template.py
====================
The downloadable, blank GL-account mapping template offered on the upload
page -- built from the user's own Mapping_Sample.xlsx label list (Koinly
Tags/Types seen across the 2025 data) so people can download it, fill in
the "Account Name" column with their real QuickBooks accounts, and upload
it back in as their mapping file.
"""
import csv
import io

# Labels in the same order as the user's Mapping_Template.csv.
LABELS = [
    # Optional overrides for the two built-in accounts -- leave blank to keep
    # "Digital Asset" / "Capital Gain/Loss".
    "Default Digital Asset Account",
    "Default Capital Gain/Loss Account",
    "Airdrop",
    "Buy",
    "Cost",
    "Crypto_Deposit",
    "Crypto_Withdrawal",
    "Exchange",
    "Fee refund",
    "Fiat_Deposit",
    "Fiat_Withdrawal",
    "From Pool",
    "Funding Fee",
    "Futures Fee",
    "Liquidity out",
    "Loan",
    "Loan Fee",
    "Loan repayment",
    "Margin loan",
    "Margin repayment",
    "Multi Trade",
    "No Label",
    "Other Fee",
    "Other Income",
    "Payment",
    "Realized gain",
    "Reward",
    "Salary",
    "Sell",
    "To Pool",
    "Transfer",
]


def build_csv():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Label", "Account Name"])
    for label in LABELS:
        w.writerow([label, ""])
    return buf.getvalue()
