"""Pandas analysis for the Honda Business Management Dashboard dataset.

Amounts are recorded transaction values in INR, not profit or confirmed receipts.
No GUI dependencies: reusable in scripts, tests, or a future notebook.
"""
from pathlib import Path
import pandas as pd

FIELDS = {
    'Rentals': ('Rental Date', 'Cost', 'Car Model', 'Customer Name'),
    'Repairs': ('DATE', 'COST', 'CAR MODEL', 'CUSTOMER NAME'),
    'Sales': ('SALE_DATE', 'SALE_PRICE', 'CAR_MODEL', 'CUSTOMER_NAME'),
}


def transaction_table(frames):
    """Normalize the three schemas to a tidy transaction table."""
    parts = []
    for category, (day, amount, model, customer) in FIELDS.items():
        frame = frames[category][[day, amount, model, customer]].copy()
        frame.columns = ['date', 'amount', 'model', 'customer']
        frame['date'] = pd.to_datetime(frame['date'], format='%Y-%m-%d', errors='raise')
        frame['amount'] = pd.to_numeric(frame['amount'], errors='raise')
        frame['category'] = category
        parts.append(frame)
    return pd.concat(parts, ignore_index=True)


def analyze(frames, year=None):
    """Return full-history or calendar-year KPIs, trends and ranked summaries.

    Monthly gaps within the selected period are filled with zeros. Year-over-year
    percentage changes and profit are intentionally not inferred from this dataset.
    """
    transactions = transaction_table(frames)
    if year is not None:
        transactions = transactions.loc[transactions['date'].dt.year.eq(int(year))].copy()
    kpis = transactions.groupby('category').agg(
        records=('amount', 'size'), recorded_amount=('amount', 'sum'),
        average_amount=('amount', 'mean'), unique_customers=('customer', 'nunique')
    ).reindex(FIELDS, fill_value=0).fillna(0)
    if year is not None:
        months = pd.date_range(f'{year}-01-01', f'{year}-12-01', freq='MS')
    elif transactions.empty:
        months = pd.DatetimeIndex([])
    else:
        months = pd.date_range(transactions['date'].min().replace(day=1),
                               transactions['date'].max().replace(day=1), freq='MS')
    monthly = (transactions.groupby([pd.Grouper(key='date', freq='MS'), 'category'])['amount']
               .sum().unstack('category').reindex(index=months, columns=FIELDS, fill_value=0).fillna(0))
    monthly.index.name = 'month'
    model_sales = (transactions.loc[transactions['category'].eq('Sales')]
                   .groupby('model').agg(sales=('amount', 'size'), recorded_amount=('amount', 'sum'))
                   .sort_values(['recorded_amount', 'model'], ascending=[False, True]))
    customers = (transactions.groupby('customer').agg(
        transactions=('amount', 'size'), recorded_amount=('amount', 'sum'))
        .sort_values(['recorded_amount', 'customer'], ascending=[False, True]).head(10))
    repair_frame = frames['Repairs'].copy()
    if year is not None:
        repair_frame = repair_frame.loc[pd.to_datetime(repair_frame['DATE']).dt.year.eq(int(year))]
    repairs = (repair_frame.groupby('ISSUES').agg(
        jobs=('COST', 'size'), recorded_amount=('COST', 'sum'), average_cost=('COST', 'mean'))
        .sort_values(['jobs', 'ISSUES'], ascending=[False, True]))
    return {'kpis': kpis, 'monthly': monthly, 'model_sales': model_sales,
            'top_customers': customers, 'repair_summary': repairs}


def export_analysis(results, directory):
    """Write the selected analysis as CSV files. No source data is changed."""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    paths = []
    for name, frame in results.items():
        path = directory / f'{name}.csv'
        frame.to_csv(path, float_format='%.2f')
        paths.append(path)
    return paths
