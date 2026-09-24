"""
CAP776 - Minor Project #1 : My Data, My Story
Name : Praveen Yadav
Reg No : 12621400
Course : MCA - D1P2633
"""

import math
import os
import openpyxl

# look for the file in the current folder first, then in a "data" folder
if os.path.exists(os.path.join("data", "12621400.xlsx")):
    FILENAME = os.path.join("data", "12621400.xlsx")
else:
    FILENAME = "12621400.xlsx"

SHEET = "Activity Log"
START_ROW = 6

# column numbers as per the excel sheet
COL_DATE = 1
COL_SLEEP = 2
COL_FITNESS = 3
COL_STUDY = 4
COL_CODING = 5
COL_CLASS = 6
COL_TOTAL_TRACKED = 9
COL_FREE = 10
COL_FEELING = 11
COL_SATISFACTION = 12
COL_ENERGY = 13

# mapping of text values to numbers (from Lists sheet)
FEELING_MAP = {"Excellent": 5, "Good": 4, "Okay": 3, "Low": 2, "Very Low": 1}
SATISFACTION_MAP = {"Very Satisfied": 5, "Satisfied": 4, "Neutral": 3,
                     "Dissatisfied": 2, "Very Dissatisfied": 1}
ENERGY_MAP = {"High": 3, "Medium": 2, "Low": 1}


def load_data():
    """
    Open the excel file ONCE and pull every needed column into a
    dictionary of lists. Returns None if the file/sheet can't be read.
    """
    try:
        wb = openpyxl.load_workbook(FILENAME, data_only=True)
    except FileNotFoundError:
        print("Error: file not found ->", FILENAME)
        print("Current folder:", os.getcwd())
        return None
    except Exception as e:
        print("Error opening workbook:", e)
        return None

    if SHEET not in wb.sheetnames:
        print("Error: sheet not found ->", SHEET)
        return None

    ws = wb[SHEET]

    data = {
        "date": [], "sleep": [], "fitness": [], "study": [], "coding": [],
        "class": [], "tracked": [], "free": [], "feeling": [],
        "satisfaction": [], "energy": [],
    }

    for row in range(START_ROW, ws.max_row + 1):
        date_val = ws.cell(row=row, column=COL_DATE).value
        if date_val is None or date_val == "":
            continue  # skip blank / unfilled rows entirely

        data["date"].append(date_val)
        data["sleep"].append(ws.cell(row=row, column=COL_SLEEP).value)
        data["fitness"].append(ws.cell(row=row, column=COL_FITNESS).value)
        data["study"].append(ws.cell(row=row, column=COL_STUDY).value)
        data["coding"].append(ws.cell(row=row, column=COL_CODING).value)
        data["class"].append(ws.cell(row=row, column=COL_CLASS).value)
        data["tracked"].append(ws.cell(row=row, column=COL_TOTAL_TRACKED).value)
        data["free"].append(ws.cell(row=row, column=COL_FREE).value)
        data["feeling"].append(ws.cell(row=row, column=COL_FEELING).value)
        data["satisfaction"].append(ws.cell(row=row, column=COL_SATISFACTION).value)
        data["energy"].append(ws.cell(row=row, column=COL_ENERGY).value)

    return data


def average(values):
    """Simple loop-based average, no numpy/pandas."""
    total = 0
    count = 0
    for v in values:
        try:
            total = total + float(v)
            count = count + 1
        except (ValueError, TypeError):
            pass
    if count == 0:
        return 0.0
    return total / count


def sum_two_lists(list_a, list_b):
    """Return element-wise sum of two equal-length lists (e.g. study+class)."""
    n = min(len(list_a), len(list_b))
    return [float(list_a[i]) + float(list_b[i]) for i in range(n)]


def map_scores(values, score_map, default):
    """Convert text values (Good/High/etc.) to numbers using a map."""
    return [score_map.get(str(v).strip(), default) for v in values]


# ---------------- Index calculations (Slide 19-27) ----------------

def calc_indices(data):
    n_days = len(data["date"])

    tpi = average(data["coding"])
    aai = average(sum_two_lists(data["study"], data["class"]))
    phai = average(data["fitness"])
    sri = average(data["sleep"])
    abi = average(data["free"])
    tui = average(data["tracked"])

    feeling_scores = map_scores(data["feeling"], FEELING_MAP, 3)
    satisfaction_scores = map_scores(data["satisfaction"], SATISFACTION_MAP, 3)
    energy_scores = map_scores(data["energy"], ENERGY_MAP, 2)

    n = min(len(feeling_scores), len(satisfaction_scores), len(energy_scores))
    ei = 0.0
    if n > 0:
        total = sum(feeling_scores[i] + satisfaction_scores[i] + energy_scores[i]
                     for i in range(n))
        ei = total / (3 * n)

    dci = (n_days / 40) * 100  # 40 expected days: 13 Aug - 21 Sept 2026

    pai = (0.15 * tpi + 0.20 * aai + 0.15 * phai + 0.20 * sri
           + 0.15 * tui + 0.10 * ei + 0.05 * dci)  # Slide 27 formula

    return {
        "n_days": n_days, "tpi": tpi, "aai": aai, "phai": phai, "sri": sri,
        "abi": abi, "tui": tui, "ei": ei, "dci": dci, "pai": pai,
        "energy_scores": energy_scores, "satisfaction_scores": satisfaction_scores,
    }


# ---------------- Correlation (Slide 28) ----------------

def pearson_r(list_a, list_b):
    n = min(len(list_a), len(list_b))
    if n < 2:
        return 0.0

    x = [float(v) for v in list_a[:n]]
    y = [float(v) for v in list_b[:n]]
    mean_x = average(x)
    mean_y = average(y)

    cov = 0
    var_x = 0
    var_y = 0
    for i in range(n):
        dx = x[i] - mean_x
        dy = y[i] - mean_y
        cov = cov + dx * dy
        var_x = var_x + dx * dx
        var_y = var_y + dy * dy

    denom = math.sqrt(var_x * var_y)
    if denom == 0:
        return 0.0
    return cov / denom


def relation_note(r):
    if abs(r) >= 0.7:
        strength = "Strong"
    elif abs(r) >= 0.3:
        strength = "Moderate"
    elif abs(r) >= 0.1:
        strength = "Weak"
    else:
        strength = "Negligible"

    direction = "positive" if r > 0 else ("negative" if r < 0 else "no")
    return f"{strength} {direction} relation (r = {r:.2f})"


# ---------------- Main ----------------

def main():
    print("=" * 60)
    print("CAP776 - Minor Project 1 : My Data, My Story")
    print("Name: Praveen Yadav | Reg No: 12621400")
    print("=" * 60)

    data = load_data()
    if not data or len(data["date"]) == 0:
        print("No data found. Check the excel file.")
        return

    n_days = len(data["date"])
    print("Valid days recorded:", n_days, "/ 40 expected")
    print("-" * 60)

    idx = calc_indices(data)

    print("Tech Productivity Index (TPI):", round(idx["tpi"], 2), "min/day")
    print("Academic Activity Index (AAI):", round(idx["aai"], 2), "min/day")
    print("Physical Activity Index (PhAI):", round(idx["phai"], 2), "min/day")
    print("Sleep & Recovery Index (SRI):", round(idx["sri"], 2), "min/day")
    print("Activity Balance Index (ABI):", round(idx["abi"], 2), "min/day")
    print("Time Utilization Index (TUI):", round(idx["tui"], 2), "min/day")
    print("Experience Index (EI):", round(idx["ei"], 2), "(scale 1-5)")
    print("Data Continuity Index (DCI):", round(idx["dci"], 2), "%")
    print("-" * 60)
    print("Personal Activity Index (PAI):", round(idx["pai"], 2))
    print("=" * 60)

    print("\nTime check: TUI + ABI =", round(idx["tui"] + idx["abi"], 2),
          "minutes (should be close to 1440)")

    print("\nRelationship Analysis:")
    r1 = pearson_r(data["coding"], idx["energy_scores"])
    r2 = pearson_r(data["sleep"], idx["energy_scores"])
    r3 = pearson_r(data["study"], idx["satisfaction_scores"])

    print("1. Coding vs Energy      ->", relation_note(r1))
    print("2. Sleep vs Energy       ->", relation_note(r2))
    print("3. Study vs Satisfaction ->", relation_note(r3))

    print("\nFindings:")
    print("- Average coding time was", round(idx["tpi"], 1), "min/day.")
    print("- Average academic time (study+class) was", round(idx["aai"], 1), "min/day.")
    print("- Average sleep was", round(idx["sri"], 1), "min/day.")
    print("- Data was recorded for", n_days, "out of 40 expected days.")
    print("=" * 60)


if __name__ == "__main__":
    main()