"""Standalone synthetic example; Python is for reproducible verification, not claimed workplace tooling."""
import csv
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
with (ROOT / 'data/synthetic_pickups.csv').open(newline='') as f:
    pickups = list(csv.DictReader(f))
with (ROOT / 'data/synthetic_daily_dates.csv').open(newline='') as f:
    dates = list(csv.DictReader(f))

def summarize(pickups, dates):
    grouped = defaultdict(list)
    for p in pickups:
        grouped[(p['Product'], p['Delivered Date'][:10])].append(p)
    result = []
    for row in dates:
        product, day = row['Product'], row['Daily Date/Time'][:10]
        matches = grouped[(product, day)]
        result.append({'Product': product, 'Date': day,
                       'Trucks': len(matches),
                       'Shipment IDs': ', '.join(p['Shipment ID'] for p in matches),
                       'Total Net Weight (LB)': sum(float(p['Net Weight (LB)']) for p in matches)})
    return result

if __name__ == '__main__':
    rows = summarize(pickups, dates)
    for p in 'ABC':
        subset = [r for r in rows if r['Product'] == p and r['Trucks']]
        print(f"Product {p}: {len(subset)} days | {sum(r['Trucks'] for r in subset)} trucks | {sum(r['Total Net Weight (LB)'] for r in subset):,.0f} LB")
    for row in rows:
        print(row)
