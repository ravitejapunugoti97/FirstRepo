import pandas as pd
import pyodbc
from azure.identity import DefaultAzureCredential


from azure.keyvault.secrets import SecretClient

# -----------------------------
# Step 1: Connect to Key Vault
# -----------------------------
import argparse

parser = argparse.ArgumentParser()

parser.add_argument("--client-name", required=True)
parser.add_argument("--persona", required=True)
parser.add_argument("--table-name", required=True)
parser.add_argument("--status-column", default="status")
parser.add_argument("--output-file", default="output.xlsx")

args = parser.parse_args()

client8name = args.client_name
persona = args.persona
table_name = args.table_name
status_column = args.status_column
excel_file = args.output_file

key_vault_url = f"https://{client8name}-eip-prod-kv.vault.azure.net/"
credential = DefaultAzureCredential()
secret_client = SecretClient(vault_url=key_vault_url, credential=credential)

# -----------------------------
# Step 2: Retrieve SQL connection string
# -----------------------------
secret_name = f"{client8name}-eip-{persona}-azuresqldb-connstring"
connection_string = secret_client.get_secret(secret_name).value

# -----------------------------
# Step 3: Parse connection string
# -----------------------------
cs_params = {}
for kv in connection_string.split(";"):
    if "=" in kv:
        k, v = kv.split("=", 1)
        cs_params[k.strip()] = v.strip()

server = cs_params["Server"].lstrip("tcp:").replace(",", ":")
database = cs_params["Initial Catalog"]
sql_user = cs_params["User ID"]
sql_pass = cs_params["Password"]

# -----------------------------
# Step 4: Connect to SQL Server
# -----------------------------
conn_str = (
    f"Driver={{ODBC Driver 17 for SQL Server}};"
    f"Server={server};"
    f"Database={database};"
    f"Uid={sql_user};"
    f"Pwd={sql_pass};"
    f"Encrypt=yes;"
    f"TrustServerCertificate=no;"
    f"Connection Timeout=30;"
)

conn = pyodbc.connect(conn_str)

# -----------------------------
# Step 5: Read table into Pandas
# -----------------------------
try:
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql(query, conn)

    if status_column not in df.columns:
        raise ValueError(f"Column '{status_column}' not found.")

    filtered_df = df[
        df[status_column].fillna("").str.upper() != "DROPPED"
    ].copy()

    filtered_df["Ignore Unit"] = ""

    filtered_df.to_excel(excel_file, index=False)

finally:
    conn.close()

print(f"Excel file saved locally as {excel_file}")
