# Imports
import pandas as pd

# BI Imports
import SQL_Invoices
import BI_ARTransactions
import SQL_Customers
import SQL_InventoryLineItems
import SQL_EstimatesCreatedOn
import SQL_Notes
import SQL_OfficeAuditTrail
import SQL_ScheduledJobs
import SQL_Memberships
import SQL_CompletedJobs
import SQL_EstimatesSoldOn
import SQL_Calls
import SQL_AppliedPayments
import SQL_AllPayments
import SQL_MarketingCampaigns
import SQL_CSR
import SQL_Rehash
import SQL_Appointments

# Standardize Files by renaming Headers and setting Data Types
def Standardize_Files(_combined_df, _ReportTypesToCombine):

    _standardized_df = pd.DataFrame()  

    match _ReportTypesToCombine:
        case "SQL_Invoices":
            _standardized_df = SQL_Invoices.Standardize_SQL_Invoices(_combined_df)            
        case "BI_ARTransactions":
            _standardized_df = BI_ARTransactions.Standardize_BI_ARTransactions(_combined_df)
        case "SQL_ARTransactions_Dec24":
            _standardized_df = BI_ARTransactions.Standardize_BI_ARTransactions(_combined_df)
        case "SQL_ARTransactions_Jan25":
            _standardized_df = BI_ARTransactions.Standardize_BI_ARTransactions(_combined_df)
        case "SQL_ARTransactions_Apr25":
            _standardized_df = BI_ARTransactions.Standardize_BI_ARTransactions(_combined_df)
        case "SQL_Customers":
            _standardized_df = SQL_Customers.Standardize_SQL_Customers(_combined_df)
        case "SQL_InventoryLineItems":
            _standardized_df = SQL_InventoryLineItems.Standardize_SQL_InventoryLineItems(_combined_df)
        case "SQL_EstimatesCreatedOn":
            _standardized_df = SQL_EstimatesCreatedOn.Standardize_SQL_EstimatesCreatedOn(_combined_df)
        case "SQL_Notes":
            _standardized_df = SQL_Notes.Standardize_SQL_Notes(_combined_df)
        case "SQL_OfficeAuditTrail":
            _standardized_df = SQL_OfficeAuditTrail.Standardize_SQL_OfficeAuditTrail(_combined_df)
        case "SQL_ScheduledJobs":
            _standardized_df = SQL_ScheduledJobs.Standardize_SQL_ScheduledJobs(_combined_df)
        case "SQL_Memberships":
            _standardized_df = SQL_Memberships.Standardize_SQL_Memberships(_combined_df)
        case "SQL_CompletedJobs":
            _standardized_df = SQL_CompletedJobs.Standardize_SQL_CompletedJobs(_combined_df)
        case "SQL_EstimatesSoldOn":
            _standardized_df = SQL_EstimatesSoldOn.Standardize_SQL_EstimatesSoldOn(_combined_df)
        case "SQL_Calls":
            _standardized_df = SQL_Calls.Standardize_SQL_Calls(_combined_df)
        case "SQL_AppliedPayments":
            _standardized_df = SQL_AppliedPayments.Standardize_SQL_AppliedPayments(_combined_df)
        case "SQL_AllPayments_":
            _standardized_df = SQL_AllPayments.Standardize_SQL_AllPayments(_combined_df)
        case "SQL_AllPaymentsExcludeDeleted":
            _standardized_df = SQL_AllPayments.Standardize_SQL_AllPayments(_combined_df)
        case "SQL_MarketingCampaigns_Today":
            _standardized_df = SQL_MarketingCampaigns.Standardize_SQL_MarketingCampaigns(_combined_df)
        case "SQL_MarketingCampaigns_Yesterday":
            _standardized_df = SQL_MarketingCampaigns.Standardize_SQL_MarketingCampaigns(_combined_df)
        case "SQL_CSR":
            _standardized_df = SQL_CSR.Standardize_SQL_CSR(_combined_df)
        case "SQL_Rehash":
            _standardized_df = SQL_Rehash.Standardize_SQL_Rehash(_combined_df)
        case "SQL_Appointments":
            _standardized_df = SQL_Appointments.Standardize_SQL_Appointments(_combined_df)
    return _standardized_df