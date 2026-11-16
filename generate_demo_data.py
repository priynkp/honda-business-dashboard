"""Create the reproducible sample dataset shipped with the project.

Run: python generate_demo_data.py
Existing sample files require --overwrite; local data/ files are never touched.
"""
import argparse
from datetime import date, timedelta
from pathlib import Path
import random

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent
SEED = 2025
MODELS = ['City', 'Amaze', 'WR-V', 'Jazz', 'Civic', 'BR-V']
SALE_PRICES = [1250000, 780000, 980000, 720000, 1850000, 1150000]
ISSUES = ['Body repair', 'Paint correction', 'Brake service', 'Engine service',
          'Electrical repair', 'Routine maintenance']

FIRST_NAMES = ['Aarav', 'Ananya', 'Priya', 'Rahul', 'Sneha', 'Vikram', 'Ishita',
               'Arjun', 'Divya', 'Karan', 'Meera', 'Nikhil', 'Pooja', 'Rohit',
               'Sanya', 'Varun', 'Aditi', 'Manish', 'Kavita', 'Suresh', 'Neha',
               'Akash', 'Riya', 'Deepak', 'Shreya', 'Anil', 'Tanvi', 'Vivek',
               'Nisha', 'Gaurav']
LAST_NAMES = ['Sharma', 'Patel', 'Verma', 'Singh', 'Gupta', 'Reddy']
CUSTOMERS = [f'{first} {last}' for last in LAST_NAMES for first in FIRST_NAMES]


def vehicle_number(number: int) -> str:
    """Format a realistic-looking plate from a 1-250 sequence value."""
    return (f'MH{10 + number % 40:02d} '
            f'{chr(65 + number % 26)}{chr(65 + (number // 26) % 26)} '
            f'{1000 + (number * 37) % 9000:04d}')


def generate_records(seed=SEED):
    """Return 600 rentals, 600 repairs, 300 sales, dated in 2024-2025."""
    rng = random.Random(seed)
    start = date(2024, 1, 1)
    result = {}
    for category, size in [('Rentals', 600), ('Repairs', 600), ('Sales', 300)]:
        rows = []
        for _ in range(size):
            day = start + timedelta(days=rng.randrange(731))
            customer = CUSTOMERS[rng.randint(1, 180) - 1]
            model = rng.choices(MODELS, weights=[30, 25, 15, 12, 8, 10])[0]
            if category == 'Rentals':
                amount = rng.randrange(20, 151) * 100
                row = [customer, model, amount, day.isoformat()]
            elif category == 'Repairs':
                issue_index = rng.randrange(len(ISSUES))
                amount = rng.randrange(15, 121) * 100 + issue_index * 250
                row = [day.isoformat(), vehicle_number(rng.randint(1, 250)),
                       customer, model, ISSUES[issue_index], amount]
            else:
                amount = round(SALE_PRICES[MODELS.index(model)] * rng.uniform(.9, 1.12) / 100) * 100
                row = [customer, model, day.isoformat(), amount]
            rows.append(row)
        date_index = {'Rentals': 3, 'Repairs': 0, 'Sales': 2}[category]
        result[category] = sorted(rows, key=lambda row: row[date_index])
    return result


def write_workbooks(directory=BASE_DIR, overwrite=False):
    # Keep generation independent of Tkinter and Matplotlib.
    schemas = {
        'Rentals': ['Customer Name', 'Car Model', 'Cost', 'Rental Date'],
        'Repairs': ['DATE', 'VEHICLE NO', 'CUSTOMER NAME', 'CAR MODEL', 'ISSUES', 'COST'],
        'Sales': ['CUSTOMER_NAME', 'CAR_MODEL', 'SALE_DATE', 'SALE_PRICE'],
    }
    directory = Path(directory)
    paths = [directory / f'{key}.xlsx' for key in schemas]
    if not overwrite and any(path.exists() for path in paths):
        raise FileExistsError('Sample files already exist. Back them up, then use --overwrite to replace them.')
    directory.mkdir(parents=True, exist_ok=True)
    records = generate_records()
    for category, headers in schemas.items():
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = category
        sheet.append(headers)
        for row in records[category]:
            sheet.append(row)
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='C80F0F')
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(
                40, max(len(str(cell.value or '')) for cell in column) + 3)
        workbook.properties.creator = ''
        workbook.properties.lastModifiedBy = ''
        workbook.save(directory / f'{category}.xlsx')
    return {category: len(rows) for category, rows in records.items()}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--overwrite', action='store_true', help='Replace the three sample workbooks only')
    args = parser.parse_args()
    print(write_workbooks(overwrite=args.overwrite))
