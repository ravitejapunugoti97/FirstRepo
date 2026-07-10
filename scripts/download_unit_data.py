import argparse
import pandas as pd
import pymssql

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

# -------------------------------------
# Parse arguments
# -------------------------------------
parser = argparse.ArgumentParser()

parser.add_argument("--client-name", required=True)
parser.add_argument("--persona", required=True)
parser.add_argument("--table-name", required=True)
parser.add_argument("--status-column", default="status")
parser.add_argument("--output-file", default="output.xlsx")

args = parser.parse_args()

client_name = args.client_name
persona = args.persona
table_name = args.table_name
status_column = args.status_column
output_file = args.output_file

# -------------------------------------
# Connect to Key Vault
# -------------------------------------
key_vault_url = f"https://{client_name}.vault.azure.net/"

credential = DefaultAzureCredential()

secret_client = SecretClient(
    vault_url=key_vault_url,
    credential=credential
)

# -------------------------------------
# Get SQL Connection String
# -------------------------------------
secret_name = f"sql-connection-string"

connection_string = secret_client.get_secret(secret_name).value

# -------------------------------------
# Parse Connection String
# -------------------------------------
params = {}

for item in connection_string.split(";"):
    if "=" in item:
        key, value = item.split("=", 1)
        params[key.strip()] = value.strip()

server = (
    params["Server"]
    .replace("tcp:", "")
    .split(",")[0]
)

database = params["Initial Catalog"]
username = params["User ID"]
password = params["Password"]

# -------------------------------------
# Connect to Azure SQL
# -------------------------------------
conn = pymssql.connect(
    server=server,
    user=username,
    password=password,
    database=database,
    port=1433
)

# -------------------------------------
# Read table
# -------------------------------------
query = f"SELECT * FROM {table_name}"

df = pd.read_sql(query, conn)

# -------------------------------------
# Filter DROPPED Units
# -------------------------------------
filtered_df = df[
    df[status_column].str.upper() != "DROPPED"
].copy()

# -------------------------------------
# Add Ignore Unit column
# -------------------------------------
filtered_df["Ignore Unit"] = ""

# -------------------------------------
# Save Excel
# -------------------------------------
filtered_df.to_excel(
    output_file,
    index=False
)

conn.close()

print(f"Excel generated successfully: {output_file}")