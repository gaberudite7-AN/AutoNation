# Imports
import pandas as pd

# Standardize Files by renaming Headers and setting Data Types
def Standardize_SQL_AllPayments(_combined_df):

    # Rename columns
    column_mapping = {
        'Payment ID':             'payment_id',
        'Payment Business Unit':  'business_unit',
        'Payment Type':           'payment_type',
        'Payment Method':         'payment_method',
        'Payment Amount':         'payment_amount',
        'Amount Allocated':       'amount_allocated',
        'Credit Remaining':       'credit_remaining',
        'Payment Status':         'payment_status',
        'Transaction Status':     'transaction_status',
        'Financing Payment':      'financing_payment',
        'Refund Method':          'refund_method',
        'Refund Reason':          'refund_reason',
        'Payment Date':           'payment_date',
        'Customer ID':            'customer_id',
        'Customer Name':          'customer_name',
        'Created By':             'created_by',
        'Customer Street':        'customer_street',
        'Customer City':          'customer_city',
        'Customer State':         'customer_state',
        'Customer Zip':           'customer_zip',
        'Customer Phone':         'customer_phone',
        'Customer Email':         'customer_email',
        'tenant':                 'tenant',
        'Batch Number':           'batch_number'
    }

    # Only include columns that exist in the DataFrame
    valid_mapping = {
        k: v for k, v in column_mapping.items() if k in _combined_df.columns
    }
    _combined_df.rename(columns=valid_mapping, inplace=True)

    # Define expected data types after renaming
    expected_dtypes = {
        'payment_id':             'string',
        'business_unit':          'string',
        'payment_type':           'string',
        'payment_method':         'string',
        'payment_amount':         'float64',
        'amount_allocated':       'float64',
        'credit_remaining':       'float64',
        'payment_status':         'string',
        'transaction_status':     'string',
        'financing_payment':      'string',
        'refund_method':          'string',
        'refund_reason':          'string',
        'payment_date':           'string',
        'customer_id':            'string',
        'customer_name':          'string',
        'created_by':             'string',
        'customer_street':        'string',
        'customer_city':          'string',
        'customer_state':         'string',
        'customer_zip':           'string',
        'customer_phone':         'string',
        'customer_email':         'string',
        'tenant':                 'string',
        'batch_number':           'string'
    }

    # Only convert types for columns that exist
    safe_dtypes = {
        col: dtype for col, dtype in expected_dtypes.items() if col in _combined_df.columns
    }
    _combined_df = _combined_df.astype(safe_dtypes)

    return _combined_df