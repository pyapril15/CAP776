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
if os.path.exists("12621400.xlsx"):
    FILENAME = "12621400.xlsx"
elif os.path.exists(os.path.join("data", "12621400.xlsx")):
    FILENAME = os.path.join("data", "12621400.xlsx")
else:
    FILENAME = "12621400.xlsx"   # will show error below if still missing

SHEET = "Activity Log"

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


def read_column(col, start_row=6):
    """Read one column from the excel sheet and return a list of values."""
    values = []
    try:
        wb = openpyxl.load_workbook(FILENAME, data_only=True)
        ws = wb[SHEET]
        for row in range(start_row, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value is not None and cell_value != "":
                values.append(cell_value)
    except FileNotFoundError:
        print("Error: file not found ->", FILENAME)
        print("Current folder:", os.getcwd())
    except Exception as e:
        print("Error while reading column", col, ":", e)
    return values


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


def days_count():
    dates = read_column(COL_DATE)
    return len(dates)


# ---------------- Index calculations (Slide 19-26) ----------------

def calc_tpi():
    coding = read_column(COL_CODING)
    return average(coding)


def calc_aai():
    study = read_column(COL_STUDY)
    cls = read_column(COL_CLASS)
    n = min(len(study), len(cls))
    combined = []
    for i in range(n):
        combined.append(float(study[i]) + float(cls[i]))
    return average(combined)


def calc_phai():
    fitness = read_column(COL_FITNESS)
    return average(fitness)


def calc_sri():
    sleep = read_column(COL_SLEEP)
    return average(sleep)


def calc_abi():
    free = read_column(COL_FREE)
    return average(free)


def calc_tui():
    tracked = read_column(COL_TOTAL_TRACKED)
    return average(tracked)


def calc_ei():
    feelings = read_column(COL_FEELING)
    satisfactions = read_column(COL_SATISFACTION)
    energies = read_column(COL_ENERGY)
    n = min(len(feelings), len(satisfactions), len(energies))
    if n == 0:
        return 0.0
    total = 0
    for i in range(n):
        f = FEELING_MAP.get(str(feelings[i]).strip(), 3)
        s = SATISFACTION_MAP.get(str(satisfactions[i]).strip(), 3)
        e = ENERGY_MAP.get(str(energies[i]).strip(), 2)
        total = total + (f + s + e)
    return total / (3 * n)


def calc_dci(expected_days=40):
    valid_days = days_count()
    return (valid_days / expected_days) * 100


def calc_pai(tpi, aai, phai, sri, tui, ei, dci):
    # Slide 27 formula (weights given by teacher)
    return (0.15 * tpi + 0.20 * aai + 0.15 * phai + 0.20 * sri
            + 0.15 * tui + 0.10 * ei + 0.05 * dci)


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

    if r > 0:
        direction = "positive"
    elif r < 0:
        direction = "negative"
    else:
        direction = "no"

    return f"{strength} {direction} relation (r = {r:.2f})"


# ---------------- Main ----------------

def main():
    print("=" * 60)
    print("CAP776 - Minor Project 1 : My Data, My Story")
    print("Name: Praveen Yadav | Reg No: 12621400")
    print("=" * 60)

    n_days = days_count()
    if n_days == 0:
        print("No data found. Check the excel file.")
        return

    print("Valid days recorded:", n_days, "/ 40 expected")
    print("-" * 60)

    tpi = calc_tpi()
    aai = calc_aai()
    phai = calc_phai()
    sri = calc_sri()
    abi = calc_abi()
    tui = calc_tui()
    ei = calc_ei()
    dci = calc_dci()
    pai = calc_pai(tpi, aai, phai, sri, tui, ei, dci)

    print("Tech Productivity Index (TPI):", round(tpi, 2), "min/day")
    print("Academic Activity Index (AAI):", round(aai, 2), "min/day")
    print("Physical Activity Index (PhAI):", round(phai, 2), "min/day")
    print("Sleep & Recovery Index (SRI):", round(sri, 2), "min/day")
    print("Activity Balance Index (ABI):", round(abi, 2), "min/day")
    print("Time Utilization Index (TUI):", round(tui, 2), "min/day")
    print("Experience Index (EI):", round(ei, 2), "(scale 1-5)")
    print("Data Continuity Index (DCI):", round(dci, 2), "%")
    print("-" * 60)
    print("Personal Activity Index (PAI):", round(pai, 2))
    print("=" * 60)

    # time budget check
    print("\nTime check: TUI + ABI =", round(tui + abi, 2),
          "minutes (should be close to 1440)")

    # relationship analysis
    print("\nRelationship Analysis:")
    coding = read_column(COL_CODING)
    sleep = read_column(COL_SLEEP)
    study = read_column(COL_STUDY)
    energy_raw = read_column(COL_ENERGY)
    satisfaction_raw = read_column(COL_SATISFACTION)

    energy_scores = [ENERGY_MAP.get(str(e).strip(), 2) for e in energy_raw]
    satisfaction_scores = [SATISFACTION_MAP.get(str(s).strip(), 3) for s in satisfaction_raw]

    r1 = pearson_r(coding, energy_scores)
    r2 = pearson_r(sleep, energy_scores)
    r3 = pearson_r(study, satisfaction_scores)

    print("1. Coding vs Energy   ->", relation_note(r1))
    print("2. Sleep vs Energy    ->", relation_note(r2))
    print("3. Study vs Satisfaction ->", relation_note(r3))

    # short findings
    print("\nFindings:")
    print("- Average coding time was", round(tpi, 1), "min/day.")
    print("- Average academic time (study+class) was", round(aai, 1), "min/day.")
    print("- Average sleep was", round(sri, 1), "min/day.")
    print("- Data was recorded for", n_days, "out of 40 expected days.")
    print("=" * 60)


if __name__ == "__main__":
    main()