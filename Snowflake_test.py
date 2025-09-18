# %%
import snowflake.connector
import pandas as pd

# Connect to Snowflake
conn = snowflake.connector.connect(
    account = "HTGXNLD-AN",
    user = "besadag@autonation.com",
    authenticator = "externalbrowser",
    role = "FINANCIAL_PLANNING_ANALYTICS_ANALYST",
    warehouse = "FINANCIAL_PLANNING_ANALYTICS_WH",
    database = "QA_AN_DW",
    schema = "BASE_URBAN_SCIENCE"
)

# Create a cursor object
cur = conn.cursor()
query = """
select top 1000 * from QA_AN_DW.BASE_URBAN_SCIENCE.DAILY_MONTH_TO_DATE_SALES where
filedate = '2025-08-11'
"""
# Execute and fetch into a DataFrame
try:
    cur.execute(query)
    df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

    # Preview the DataFrame
    print(df.head())


finally:
    cur.close()
    conn.close()