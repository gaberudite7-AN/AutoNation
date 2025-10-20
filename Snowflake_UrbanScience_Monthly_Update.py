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
    database = "WORKSPACES",
    schema = "FINANCIAL_PLANNING_ANALYTICS"
)

# We will be pulling the latest data from snowflake for Cube Historical and YTD. 
# Snowflake query is dynamic and will update latest daily file and append to historical while making the necessary transformations.
cur = conn.cursor()
YTD_query = """
select * from workspaces.financial_planning_analytics.Cube_YTD
"""

Historical_query = """
select * from workspaces.financial_planning_analytics.CUBE_HISTORICAL_VIEW
"""

# Execute and fetch into a DataFrame
try:
    cur.execute(YTD_query)
    YTD_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

    cur.execute(Historical_query)
    Historical_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

    # Preview the DataFrame
    YTD_df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Excels\YTD_df.csv', index=False)
    Historical_df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Excels\Historical_df.csv', index=False)

finally:
    cur.close()
    conn.close()