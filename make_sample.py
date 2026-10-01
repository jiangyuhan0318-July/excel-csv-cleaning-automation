"""Create an original messy order sample in both CSV and Excel formats."""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from pipeline import COLUMNS


def generate(root):
    root = Path(root)
    if root.exists():
        raise FileExistsError(root)
    rows = [
        [' o001 ', ' Alice   Chen ', ' ALICE@EXAMPLE.COM ', ' office supplies ', '2026/09/01', '$120.00', ' PAID '],
        ['O002', 'Bob Smith', 'bob@example.com', 'TECH', '09/02/2026', '250', 'completed'],
        ['O003', 'Carla Jones', '', 'Furniture', '03-Sep-2026', '450.50', 'Pending'],
        ['o004', 'David Lee', 'david@example.com', 'office', '2026-09-04', '75.5', 'complete'],
        ['O005', 'Eva Patel', 'eva@example.com', 'technology', '2026/09/05', '1,200.00', 'paid'],
        ['O006', 'Frank Wu', 'frank@example.com', 'OFFICE', '09/06/2026', '0', 'canceled'],
        ['O007', 'Grace Kim', 'grace@example.com', 'furniture', '07-Sep-2026', '$320', 'pending'],
        ['O008', 'Henry Park', '', 'Tech', '2026-09-08', '89.99', 'Paid'],
        ['O009', 'Ivy Brown', 'ivy@example.com', 'Office Supplies', '2026/09/09', '45', 'paid'],
        ['O010', 'Jack Lin', 'jack@example.com', 'Furniture', '09/10/2026', '600.00', 'Completed'],
        ['O001', 'Alice Chen', 'alice@example.com', 'Office', '2026-09-01', '120', 'Paid'],
        ['O002', 'Bob Smith', 'bob@example.com', 'Technology', '2026-09-02', '$250.00', 'Paid'],
        ['O011', '', 'kate@example.com', 'Office', '2026-09-11', '15', 'Paid'],
        ['O012', 'Leo Sun', 'leo@example.com', 'Tech', '2026-09-12', '', 'Paid'],
        ['O013', 'Mia Hall', 'mia@example.com', 'Furniture', '2026-02-30', '80', 'Pending'],
        ['O014', 'Noah King', 'noah@example.com', 'Mystery', '2026-09-14', '30', 'Paid'],
        ['O015', 'Olivia Fox', 'olivia@example.com', 'Office', '2026-09-15', '-20', 'Paid'],
        ['O016', 'Paul Reed', 'not-an-email', 'Tech', '2026-09-16', '100', 'Paid'],
        ['O017', 'Quinn Ray', 'quinn@example.com', 'Office', '2026-09-17', '40', 'Paid'],
        ['O017', 'Quinn Ray', 'quinn@example.com', 'Office', '2026-09-17', '90', 'Paid'],
    ]
    root.mkdir(parents=True)
    pd.DataFrame(rows, columns=COLUMNS).to_csv(root/'messy_orders.csv', index=False)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Orders'
    ws.append(COLUMNS)
    for row in rows:
        ws.append(row)
    wb.save(root/'messy_orders.xlsx')


if __name__ == '__main__':
    generate(Path(__file__).parent/'sample_input')
