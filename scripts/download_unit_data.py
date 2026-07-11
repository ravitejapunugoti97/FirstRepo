import argparse
import pandas as pd
import pymssql

from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
from azure.storage.filedatalake import DataLakeServiceClient

# -------------------------------------
# Parse arguments
# -------------------------------------
parser = argparse.ArgumentParser()

parser.add_argument("--client-name", required=True)
parser.add_argument("--persona", required=True)
parser.add_argument("--table-name", required=True)
parser.add_argument("--status-column", default="status")
parser.add_argument("--output-file", default="output.xlsx")
parser.add_argument("--storage-account", required=True)
parser.add_argument("--filesystem", required=True)
directory = client_name

args = parser.parse_args()

client_name = args.client_name
persona = args.persona
table_name = args.table_name
status_column = args.status_column
output_file = args.output_file
storage_account = args.storage_account
filesystem = args.filesystem
directory = client_name

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
try:
    # -------------------------------------
    # Read table
    # -------------------------------------
    query = f"SELECT * FROM {persona}.{table_name}"

    print(f"Executing Query: {query}")

    df = pd.read_sql(query, conn)

    print(f"Total rows fetched: {len(df)}")

    if status_column not in df.columns:
        raise ValueError(
            f"Column '{status_column}' not found."
        )

    # -------------------------------------
    # Filter DROPPED Units
    # -------------------------------------
    filtered_df = df[
        df[status_column].fillna("").str.upper() != "DROPPED"
    ].copy()

    print(f"Rows after filtering: {len(filtered_df)}")

    # -------------------------------------
    # Add Ignore Unit column
    # -------------------------------------
    filtered_df["Ignore Unit"] = ""

    # -------------------------------------
    # Save Excel
    # -------------------------------------
    filtered_df.to_excel(output_file, index=False)

    print(f"Excel generated successfully: {output_file}")

    # -------------------------------------
    # Upload Excel to ADLS Gen2
    # -------------------------------------
    account_url = f"https://{storage_account}.dfs.core.windows.net"

    service_client = DataLakeServiceClient(
        account_url=account_url,
        credential=credential
    )

    filesystem_client = service_client.get_file_system_client(
        filesystem
    )

    # Create folder using client name
    directory_client = filesystem_client.get_directory_client(
        directory
    )

    file_client = directory_client.get_file_client(
        output_file
    )

    with open(output_file, "rb") as file:
        file_data = file.read()

    file_client.upload_data(
        file_data,
        overwrite=True
    )

    print(
        f"Excel uploaded successfully to "
        f"{filesystem}/{client_name}/{output_file}"
    )

except Exception as e:
    print(f"Error: {e}")
    raise

finally:
    conn.close()