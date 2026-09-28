"""
mapping_template.py
====================
The downloadable GL-account mapping template offered on the upload page.
It comes PRE-FILLED with a default mapping (Label -> suggested QuickBooks
account). It is only a starting point: users should change each Account
Name to the EXACT account name in their own QuickBooks chart of accounts
before uploading it back as their mapping file.
"""
import csv
import io

# Default mapping (Label, suggested Account Name), in template order.
DEFAULT_MAPPING = [
    # Optional overrides for the two built-in accounts -- leave blank to keep
    # "Digital Asset" / "Capital Gain/Loss".
    ("Default Digital Asset Account", "Digital Assets"),
    ("Default Capital Gain/Loss Account", "Cryptocurrency Capital Gain/Loss"),
    ("Airdrop", "Crypto Consulting Income"),
    ("Buy", "Digital Assets"),
    ("Cost", "Crypto fees"),
    ("Crypto_Deposit", "Digital Assets"),
    ("Crypto_Withdrawal", "Digital Assets"),
    ("Exchange", "Digital Assets"),
    ("Fee refund", "Crypto fees"),
    ("Fiat_Deposit", "Capital Contribution"),
    ("Fiat_Withdrawal", "Bank"),
    ("From Pool", "Digital Assets"),
    ("Funding Fee", "Ask Client"),
    ("Futures Fee", "Ask Client"),
    ("Liquidity out", "Digital Assets"),
    ("Loan", "Capital Contribution"),
    ("Loan Fee", "Crypto fees"),
    ("Loan repayment", "Distributions"),
    ("Margin loan", "Loan to Others"),
    ("Margin repayment", "Loan to Others"),
    ("Multi Trade", "Digital Assets"),
    ("No Label", "Aks"),
    ("Other Fee", "Contractors"),
    ("Other Income", "Income"),
    ("Payment", "Contractors"),
    ("Realized gain", "Cryptocurrency Capital Gain/Loss"),
    ("Reward", "Rewards"),
    ("Salary", "Crypto Consulting Income"),
    ("Sell", "Digital Assets"),
    ("To Pool", "Digital Assets"),
    ("Transfer", "Digital Assets"),
]

LABELS = [label for label, _ in DEFAULT_MAPPING]


def build_csv():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Label", "Account Name"])
    for label, acct in DEFAULT_MAPPING:
        w.writerow([label, acct])
    return buf.getvalue()
