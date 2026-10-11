"""Generate a synthetic dataset for a small grocery store (3 months).

All data is fictional. Run it once to (re)create the CSV files in ./data.

    python generate_sample_data.py
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 42
START = date(2026, 7, 1)
END = date(2026, 9, 30)
DATA_DIR = Path(__file__).parent / "data"

# product_id, name, category, unit_cost, unit_price, avg units/day, stock_on_hand, supplier
PRODUCTS = [
    ("P01", "Cola 600ml",         "Beverages", 0.60, 1.00, 26, 40,  "Beverage Wholesale"),
    ("P02", "Orange Soda 600ml",  "Beverages", 0.55, 0.95, 18, 120, "Beverage Wholesale"),
    ("P03", "Bottled Water 1L",   "Beverages", 0.30, 0.60, 30, 55,  "Beverage Wholesale"),
    ("P04", "Fruit Juice 1L",     "Beverages", 0.90, 1.50, 12, 70,  "Beverage Wholesale"),
    ("P05", "Potato Chips",       "Snacks",    0.80, 1.30, 20, 130, "Snack Distributor"),
    ("P06", "Salted Peanuts",     "Snacks",    0.50, 0.90, 14, 90,  "Snack Distributor"),
    ("P07", "Cookies",            "Snacks",    0.70, 1.20, 16, 110, "Snack Distributor"),
    ("P08", "Chocolate Bar",      "Snacks",    0.45, 0.85, 22, 150, "Snack Distributor"),
    ("P09", "Whole Milk 1L",      "Dairy",     0.85, 1.20, 24, 35,  "Dairy Co-op"),
    ("P10", "Yogurt 1kg",         "Dairy",     1.40, 2.20, 10, 45,  "Dairy Co-op"),
    ("P11", "Cheese 400g",        "Dairy",     2.60, 3.80, 7,  30,  "Dairy Co-op"),
    ("P12", "Butter 200g",        "Dairy",     1.30, 1.90, 6,  28,  "Dairy Co-op"),
    ("P13", "Dish Soap",          "Cleaning",  1.10, 1.90, 9,  12,  "Home Supplies Inc."),
    ("P14", "Laundry Detergent",  "Cleaning",  3.50, 5.20, 5,  26,  "Home Supplies Inc."),
    ("P15", "Bleach 1L",          "Cleaning",  0.90, 1.50, 8,  48,  "Home Supplies Inc."),
    ("P16", "Paper Towels",       "Cleaning",  1.60, 2.50, 7,  36,  "Home Supplies Inc."),
]

MONTH_TREND = {7: 1.00, 8: 1.04, 9: 1.08}   # slow growth over the quarter
WEEKDAY_FACTOR = {4: 1.15, 5: 1.35, 6: 1.30}  # Fri, Sat, Sun sell more
LEAD_TIME_DAYS = 3  # reorder level = demand during the supplier lead time


def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def make_sales(rng):
    rows = []
    for d in daterange(START, END):
        factor = MONTH_TREND[d.month] * WEEKDAY_FACTOR.get(d.weekday(), 1.0)
        for pid, name, cat, cost, price, avg, *_ in PRODUCTS:
            units = round(rng.gauss(avg * factor, avg * 0.25))
            if units > 0:
                rows.append([d.isoformat(), pid, name, cat, units, price])
    return rows


def make_expenses(rng):
    rows = []

    def add(d, cat, desc, amount):
        rows.append([d.isoformat(), cat, desc, f"{amount:.2f}"])

    for month in (7, 8, 9):
        add(date(2026, month, 1), "Rent", "Store rent", 900)
        add(date(2026, month, 15), "Payroll", "Payroll (first half)", 700)
        add(date(2026, month, 28), "Payroll", "Payroll (second half)", 700)
        add(date(2026, month, 5), "Utilities", "Electricity", round(rng.uniform(180, 235), 2))
        add(date(2026, month, 8), "Utilities", "Water", round(rng.uniform(40, 55), 2))
        add(date(2026, month, 10), "Utilities", "Internet", 45)
        add(date(2026, month, 12), "Marketing", "Flyers and social media ads", round(rng.uniform(80, 120), 2))
        for _ in range(rng.randint(4, 6)):
            day = date(2026, month, rng.randint(1, 28))
            desc = rng.choice(["Shopping bags", "Receipt paper", "Cleaning supplies", "Price labels"])
            add(day, "Supplies", desc, round(rng.uniform(15, 70), 2))
        add(date(2026, month, rng.randint(3, 25)), "Maintenance", "Minor repairs", round(rng.uniform(20, 40), 2))
    add(date(2026, 8, 19), "Maintenance", "Refrigerator repair", 160)
    rows.sort(key=lambda r: r[0])
    return rows


def make_inventory():
    rows = []
    for pid, name, cat, cost, price, avg, stock, supplier in PRODUCTS:
        reorder_level = round(avg * LEAD_TIME_DAYS)
        rows.append([pid, name, cat, cost, price, stock, reorder_level, supplier])
    return rows


def write_csv(filename, header, rows):
    path = DATA_DIR / filename
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    print(f"Wrote {len(rows):>5} rows -> {path.relative_to(Path(__file__).parent)}")


def main():
    rng = random.Random(SEED)
    DATA_DIR.mkdir(exist_ok=True)
    write_csv("sample_sales.csv",
              ["date", "product_id", "product", "category", "units", "unit_price"],
              make_sales(rng))
    write_csv("sample_expenses.csv",
              ["date", "category", "description", "amount"],
              make_expenses(rng))
    write_csv("sample_inventory.csv",
              ["product_id", "product", "category", "unit_cost", "unit_price",
               "stock_on_hand", "reorder_level", "supplier"],
              make_inventory())


if __name__ == "__main__":
    main()
