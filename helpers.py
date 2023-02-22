import pandas as pd

def clean_auc(df):
    df = df.copy()
    neg_spike = df['Spike endpoint'].astype(str).str.strip().str.upper().str[:2] == "NE"
    neg_auc = df['AUC'] == 0.
    df.loc[neg_spike | neg_auc, 'AUC'] = 1.
    return df['AUC']

def clean_sample_id(df):
    return df['Sample ID'].astype(str).str.strip().str.upper()

def clean_research(df):
    return (df.assign(sample_id=clean_sample_id,
                      AUC=clean_auc)
              .dropna(subset=['AUC'])
              .query("AUC not in ['-']")
              .loc[:, ['sample_id', 'Spike endpoint', 'AUC']])

def try_date(potential_date):
    return pd.to_datetime(potential_date, errors='coerce').date()

def try_datediff(start_date, end_date):
    try:
        return int((end_date - start_date).days)
    except:
        return 'N/A'

def permissive_datemax(date_list, comp_date):
    placeholder = pd.to_datetime('1.1.1950').date()
    max_date = placeholder
    for date_ in date_list:
        date_ = pd.to_datetime(date_, errors='coerce').date()
        if not pd.isna(date_) and date_ > max_date and date_ < comp_date:
            max_date = date_
    if max_date > placeholder:
        return max_date
