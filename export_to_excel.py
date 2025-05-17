import urllib.parse
import pandas as pd
from sqlalchemy import create_engine
import urllib

# components
server= 'Your_server_name'
driver = "{ODBC Driver 17 for SQL Server}"
db='Your_Database_name'
table='Table_name'
path=r'path\to\your\excel_file\demodata.xlsx'

# connection
conn = urllib.parse.quote_plus(
    f"DRIVER={driver};"
    f"SERVER={server};"
    f"DATABASE={db};"
    "Trusted_Connection=yes;"
)

# engine using sqlalchemy
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={conn}")

# query and export data
try:
    query = f''' select * from {table} '''
    df = pd.read_sql(query,engine)
    df.to_excel(path,index=False,engine='openpyxl')
    print("Successfully exported files!")
except Exception as e:
    print(f"{e}: File export failed!")