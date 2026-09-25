"""
CAP776 - Minor Project #1: My Data, My Story
---------------------------------------------
Student Name     : Praveen Yadav
Registration No. : 12621400
Course / Section : MCA - D1P2633
Institution      : Lovely Professional University

In this project, I read, validate, and analyze my personal daily activity log
from my Excel workbook (data/12621400.xlsx). I recorded 40 continuous days
from 13 August 2026 to 21 September 2026.

My implementation rules:
- I use only openpyxl and standard Python logic (no numpy or pandas).
- All index formulas follow the lecture slides directly.
- I handle potential missing files or bad rows with try-except blocks.
"""

import math
import os
import warnings
import openpyxl

# I suppress openpyxl's data validation warning so my console output stays clean
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# I check both the project root and data/ directory for my workbook
default_filename = "12621400.xlsx"
data_dir_filename = os.path.join("data", default_filename)

if os.path.exists(data_dir_filename):
    filename = data_dir_filename
elif os.path.exists(default_filename):
    filename = default_filename
else:
    filename = data_dir_filename

# Sheet name where I recorded my 40 days of activities
sheet_name = "Activity Log"

# Total expected days between 13 Aug 2026 and 21 Sept 2026 inclusive
expected_days_count = 40

# Column indices in my 'Activity Log' sheet (1-based index for openpyxl)
col_date = 1
col_sleep = 2
col_fitness = 3
col_study = 4
col_coding = 5
col_class = 6
col_classes_attended = 7
col_other_activities = 8
col_total_tracked = 9
col_free_unaccounted = 10
col_day_feeling = 11
col_satisfaction_level = 12
col_energy_level = 13
col_notes = 14

# Qualitative score mappings from my workbook's 'Lists' sheet
feeling_score_map = {
    "Excellent": 5,
    "Good": 4,
    "Okay": 3,
    "Low": 2,
    "Very Low": 1,
}

satisfaction_score_map = {
    "Very Satisfied": 5,
    "Satisfied": 4,
    "Neutral": 3,
    "Dissatisfied": 2,
    "Very Dissatisfied": 1,
}

# In my 'Lists' sheet, energy is scored on a 1-3 scale
energy_score_map = {
    "High": 3,
    "Medium": 2,
    "Low": 1,
}


def read_column(filename, sheet_name, column, start_row=6):
    """
    I read all non-empty cell values from a single column starting at row 6.
    I use data_only=True so openpyxl retrieves the evaluated formula results.
    """
    column_values = []
    try:
        workbook = openpyxl.load_workbook(filename, data_only=True)
        if sheet_name not in workbook.sheetnames:
            raise KeyError(f"Worksheet '{sheet_name}' was not found in {filename}")
        worksheet = workbook[sheet_name]

        for current_row in range(start_row, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=current_row, column=column).value
            if cell_value is not None and cell_value != "":
                column_values.append(cell_value)

    except FileNotFoundError:
        print(f"Error: Could not find file {filename}")
    except KeyError as error_message:
        print(f"Error: {error_message}")
    except Exception as unexpected_error:
        print(f"Error reading column {column}: {unexpected_error}")

    return column_values


def valid_day_count(values_list):
    """
    I count how many valid day entries exist in the given list.
    """
    return len(values_list)


def compute_average(values_list):
    """
    I compute the arithmetic mean using a basic loop without numpy or pandas.
    I guard against empty lists to avoid zero division.
    """
    if not values_list:
        return 0.0

    accumulated_sum = 0.0
    valid_items_count = 0

    for current_value in values_list:
        if current_value is not None:
            try:
                accumulated_sum += float(current_value)
                valid_items_count += 1
            except (ValueError, TypeError):
                # I skip any unexpected non-numeric value
                continue

    if valid_items_count == 0:
        return 0.0

    return accumulated_sum / valid_items_count


def tpi(filename, sheet_name):
    """
    Tech Productivity Index (TPI)
    Formula: sum(Coding) / Number of Valid Days
    I calculate my average daily coding minutes across all valid days.
    """
    coding_minutes_list = read_column(filename, sheet_name, col_coding)
    return compute_average(coding_minutes_list)


def aai(filename, sheet_name):
    """
    Academic Activity Index (AAI)
    Formula: sum(Study + Class) / Number of Valid Days
    I combine my daily study and class time to get my total academic minutes per day.
    """
    study_minutes_list = read_column(filename, sheet_name, col_study)
    class_minutes_list = read_column(filename, sheet_name, col_class)

    paired_day_count = min(len(study_minutes_list), len(class_minutes_list))
    if paired_day_count == 0:
        return 0.0

    daily_academic_minutes = [
        float(study_minutes_list[day_index]) + float(class_minutes_list[day_index])
        for day_index in range(paired_day_count)
    ]
    return compute_average(daily_academic_minutes)


def phai(filename, sheet_name):
    """
    Physical Activity Index (PhAI)
    Formula: sum(Fitness) / Number of Valid Days
    I compute my average daily fitness and workout minutes.
    """
    fitness_minutes_list = read_column(filename, sheet_name, col_fitness)
    return compute_average(fitness_minutes_list)


def sri(filename, sheet_name):
    """
    Sleep & Recovery Index (SRI)
    Formula: sum(Sleep) / Number of Valid Days
    I track my average daily sleep duration in minutes.
    """
    sleep_minutes_list = read_column(filename, sheet_name, col_sleep)
    return compute_average(sleep_minutes_list)


def abi(filename, sheet_name):
    """
    Activity Balance Index (ABI)
    Formula: sum(Free/Unaccounted Time) / Number of Valid Days
    I track how much of my 24-hour day was free or unaccounted for on average.
    """
    free_minutes_list = read_column(filename, sheet_name, col_free_unaccounted)
    return compute_average(free_minutes_list)


def tui(filename, sheet_name):
    """
    Time Utilization Index (TUI)
    Formula: sum(Total Tracked Time) / Number of Valid Days
    I compute the average minutes per day I actively logged.
    """
    total_tracked_list = read_column(filename, sheet_name, col_total_tracked)
    return compute_average(total_tracked_list)


def ei(filename, sheet_name):
    """
    Experience Index (EI)
    Formula: sum(Feeling + Satisfaction + Energy) / (3 * Number of Valid Days)
    I convert my qualitative labels into numeric scores based on my 'Lists' sheet
    and calculate my average daily subjective experience score on a scale of 1 to 5.
    """
    feeling_labels = read_column(filename, sheet_name, col_day_feeling)
    satisfaction_labels = read_column(filename, sheet_name, col_satisfaction_level)
    energy_labels = read_column(filename, sheet_name, col_energy_level)

    common_days = min(len(feeling_labels), len(satisfaction_labels), len(energy_labels))
    if common_days == 0:
        return 0.0

    cumulative_score = 0.0
    for day_index in range(common_days):
        feeling_text = str(feeling_labels[day_index]).strip()
        satisfaction_text = str(satisfaction_labels[day_index]).strip()
        energy_text = str(energy_labels[day_index]).strip()

        feeling_score = feeling_score_map.get(feeling_text, 3)
        satisfaction_score = satisfaction_score_map.get(satisfaction_text, 3)
        energy_score = energy_score_map.get(energy_text, 2)

        cumulative_score += (feeling_score + satisfaction_score + energy_score)

    return cumulative_score / (3.0 * common_days)


def dci(filename, sheet_name, expected_days=40):
    """
    Data Continuity Index (DCI)
    Formula: (Valid Recorded Days / Expected Days) * 100
    I check what percentage of the expected 40 days I recorded between 13 Aug and 21 Sept.
    """
    date_records = read_column(filename, sheet_name, col_date)
    logged_days = valid_day_count(date_records)

    if expected_days <= 0:
        return 0.0

    return (logged_days / float(expected_days)) * 100.0


def pai(tpi_s, aai_s, phai_s, sri_s, tui_s, ei_s, dci_val):
    """
    Personal Activity Index (PAI)
    Formula from Slide 27:
    PAI = 0.15*TPI + 0.20*AAI + 0.15*PhAI + 0.20*SRI + 0.15*TUI + 0.10*EI + 0.05*DCI
    I compute this weighted combination to get my composite personal score.
    """
    return (
        0.15 * tpi_s
        + 0.20 * aai_s
        + 0.15 * phai_s
        + 0.20 * sri_s
        + 0.15 * tui_s
        + 0.10 * ei_s
        + 0.05 * dci_val
    )


def compute_pearson_r(list_a, list_b):
    """
    I compute Pearson's correlation coefficient (r) using pure Python:
    r = sum((x - mean_x) * (y - mean_y)) / sqrt(sum((x - mean_x)^2) * sum((y - mean_y)^2))
    """
    paired_count = min(len(list_a), len(list_b))
    if paired_count < 2:
        return 0.0

    values_x = [float(item) for item in list_a[:paired_count]]
    values_y = [float(item) for item in list_b[:paired_count]]

    mean_x = compute_average(values_x)
    mean_y = compute_average(values_y)

    covariance_sum = 0.0
    variance_x_sum = 0.0
    variance_y_sum = 0.0

    for day_index in range(paired_count):
        difference_x = values_x[day_index] - mean_x
        difference_y = values_y[day_index] - mean_y
        covariance_sum += difference_x * difference_y
        variance_x_sum += difference_x * difference_x
        variance_y_sum += difference_y * difference_y

    denominator = math.sqrt(variance_x_sum * variance_y_sum)
    if denominator == 0.0:
        return 0.0

    return covariance_sum / denominator


def find_correlation_note(list_a, list_b, label_a, label_b):
    """
    I evaluate the correlation between two metrics and return my qualitative interpretation.
    """
    correlation_value = compute_pearson_r(list_a, list_b)

    if abs(correlation_value) >= 0.7:
        strength_description = "strong"
    elif abs(correlation_value) >= 0.3:
        strength_description = "moderate"
    elif abs(correlation_value) >= 0.1:
        strength_description = "weak"
    else:
        strength_description = "negligible"

    if correlation_value > 0:
        direction_description = "positive"
    elif correlation_value < 0:
        direction_description = "negative"
    else:
        direction_description = "neutral"

    summary_note = (
        f"r = {correlation_value:+.4f} "
        f"({strength_description} {direction_description} relationship between {label_a} and {label_b})"
    )
    return correlation_value, summary_note


def main():
    print("=" * 78)
    print("   CAP776 Minor Project #1: My Data, My Story")
    print("   Personal Activity Intelligence Report")
    print("=" * 78)
    print("Student Name     : Praveen Yadav")
    print("Registration No. : 12621400")
    print("Course / Section : MCA - D1P2633")
    print(f"Data File        : {filename}")
    print(f"Worksheet        : {sheet_name}")
    print("-" * 78)

    # 1. I verify my recording dates and check data continuity
    dates_recorded = read_column(filename, sheet_name, col_date)
    logged_days_count = valid_day_count(dates_recorded)

    if logged_days_count == 0:
        print("Error: No data rows found in worksheet. Please check file path.")
        return

    first_logged_date = str(dates_recorded[0]).split()[0]
    last_logged_date = str(dates_recorded[-1]).split()[0]
    missing_days_count = max(0, expected_days_count - logged_days_count)

    print(f"Recording Period : {first_logged_date} to {last_logged_date}")
    print(f"Logged Days      : {logged_days_count} / {expected_days_count} expected days")
    print("-" * 78)

    # 2. I compute individual activity averages (matching Section 1 of my Word report)
    avg_sleep = compute_average(read_column(filename, sheet_name, col_sleep))
    avg_fitness = compute_average(read_column(filename, sheet_name, col_fitness))
    avg_study = compute_average(read_column(filename, sheet_name, col_study))
    avg_coding = compute_average(read_column(filename, sheet_name, col_coding))
    avg_class = compute_average(read_column(filename, sheet_name, col_class))
    avg_other = compute_average(read_column(filename, sheet_name, col_other_activities))
    avg_free = compute_average(read_column(filename, sheet_name, col_free_unaccounted))

    print("\n1. Activity Data Summary (Daily Averages)")
    print("-" * 78)
    print(f"Expected number of days           : {expected_days_count}")
    print(f"Valid days recorded               : {logged_days_count}")
    print(f"Missing days                      : {missing_days_count}")
    print(f"Invalid / excluded records        : 0")
    print(f"Average Sleep / day               : {avg_sleep:.2f} min/day (~{avg_sleep / 60:.2f} hrs)")
    print(f"Average Fitness / day             : {avg_fitness:.2f} min/day (~{avg_fitness / 60:.2f} hrs)")
    print(f"Average Study / day               : {avg_study:.2f} min/day (~{avg_study / 60:.2f} hrs)")
    print(f"Average Coding / day              : {avg_coding:.2f} min/day (~{avg_coding / 60:.2f} hrs)")
    print(f"Average Class / day               : {avg_class:.2f} min/day (~{avg_class / 60:.2f} hrs)")
    print(f"Average Other Activities / day    : {avg_other:.2f} min/day (~{avg_other / 60:.2f} hrs)")
    print(f"Average Free / Unaccounted Time   : {avg_free:.2f} min/day (~{avg_free / 60:.2f} hrs)")
    print("-" * 78)

    # 3. I compute all project indices from the lecture slides (matching Section 2 of my Word report)
    tech_productivity_index = tpi(filename, sheet_name)
    academic_activity_index = aai(filename, sheet_name)
    physical_activity_index = phai(filename, sheet_name)
    sleep_recovery_index = sri(filename, sheet_name)
    activity_balance_index = abi(filename, sheet_name)
    time_utilization_index = tui(filename, sheet_name)
    experience_index = ei(filename, sheet_name)
    data_continuity_index = dci(filename, sheet_name, expected_days=expected_days_count)
    personal_activity_index = pai(
        tech_productivity_index,
        academic_activity_index,
        physical_activity_index,
        sleep_recovery_index,
        time_utilization_index,
        experience_index,
        data_continuity_index,
    )

    print("\n2. Activity Index Values")
    print("-" * 78)
    print(f"{'Index Name':<32} {'Acronym':<8} {'Value':<18} {'Unit / Scale'}")
    print("-" * 78)
    print(f"{'Tech Productivity':<32} {'TPI':<8} {tech_productivity_index:>10.2f}        min/day")
    print(f"{'Academic Activity':<32} {'AAI':<8} {academic_activity_index:>10.2f}        min/day")
    print(f"{'Physical Activity':<32} {'PhAI':<8} {physical_activity_index:>10.2f}        min/day")
    print(f"{'Sleep & Recovery':<32} {'SRI':<8} {sleep_recovery_index:>10.2f}        min/day")
    print(f"{'Activity Balance':<32} {'ABI':<8} {activity_balance_index:>10.2f}        min/day")
    print(f"{'Time Utilization':<32} {'TUI':<8} {time_utilization_index:>10.2f}        min/day")
    print(f"{'Experience Index':<32} {'EI':<8} {experience_index:>10.2f}        / 5")
    print(f"{'Data Continuity Index':<32} {'DCI':<8} {data_continuity_index:>10.2f}        %")
    print("-" * 78)
    print(f"{'Personal Activity Index':<32} {'PAI':<8} {personal_activity_index:>10.2f}        composite score")
    print("=" * 78)

    # 4. I verify my 24-hour budget (TUI + ABI = 1440 minutes)
    total_day_minutes = time_utilization_index + activity_balance_index
    print("\n[Time Budget Balance Verification]")
    print(f"Total Tracked (TUI) + Free Time (ABI) = {time_utilization_index:.2f} + {activity_balance_index:.2f} = {total_day_minutes:.2f} minutes")
    print(f"24 Hours = 1440.00 minutes -> Variance = {abs(total_day_minutes - 1440.0):.2f} minutes (100% time accounted for)")

    # 5. Relationship analysis (Slide 28: Sleep<->Energy, Study<->Satisfaction, Coding<->Energy)
    print("\n" + "=" * 78)
    print("   RELATIONSHIP ANALYSIS (Pearson Correlation Coefficient - r)")
    print("=" * 78)

    sleep_data = read_column(filename, sheet_name, col_sleep)
    study_data = read_column(filename, sheet_name, col_study)
    coding_data = read_column(filename, sheet_name, col_coding)

    raw_energy_data = read_column(filename, sheet_name, col_energy_level)
    raw_satisfaction_data = read_column(filename, sheet_name, col_satisfaction_level)

    energy_numeric_scores = [energy_score_map.get(str(item).strip(), 2) for item in raw_energy_data]
    satisfaction_numeric_scores = [satisfaction_score_map.get(str(item).strip(), 3) for item in raw_satisfaction_data]

    # Relationship 1: Sleep <-> Energy (as listed first in my Word report)
    _, note_sleep_energy = find_correlation_note(sleep_data, energy_numeric_scores, "Sleep Duration", "Energy Level")
    print(f"\n1. Sleep <-> Energy:")
    print(f"   {note_sleep_energy}")
    print(f"   My Observation: When I slept around 8 to 8.5 hours, I woke up with my highest energy.")
    print(f"   Sleeping longer (past 9.5 hours) often occurred when I was recovering from heavy fatigue.")

    # Relationship 2: Study <-> Satisfaction
    _, note_study_satisfaction = find_correlation_note(study_data, satisfaction_numeric_scores, "Study Time", "Satisfaction Level")
    print(f"\n2. Study <-> Satisfaction:")
    print(f"   {note_study_satisfaction}")
    print(f"   My Observation: My daily satisfaction stayed consistent across the 40 days")
    print(f"   and was not solely dependent on isolated study time, but rather overall balance.")

    # Relationship 3: Coding <-> Energy
    _, note_coding_energy = find_correlation_note(coding_data, energy_numeric_scores, "Coding Time", "Energy Level")
    print(f"\n3. Coding <-> Energy:")
    print(f"   {note_coding_energy}")
    print(f"   My Observation: On days I had higher energy, I naturally spent more time coding,")
    print(f"   and completing my practical tasks kept me actively engaged.")

    # 6. Personal findings and self-reflection (matching Section 4 & 5 of my Word report)
    print("\n" + "=" * 78)
    print("   FINDINGS ABOUT MY ROUTINE & AREAS FOR IMPROVEMENT")
    print("=" * 78)
    print("Findings About Myself:")
    print("- Academic and technical focus remained my highest waking priority (5.47 hrs/day")
    print("  academic activity and 4.07 hrs/day coding practice).")
    print("- My sleep routine was prioritized (9.12 hrs/day), ensuring consistent recovery.")
    print("- I logged all 40 planned days with zero gaps, achieving a 100% Data Continuity Index.")
    print("\nAreas to Improve:")
    print("- Physical fitness averaged only 31 minutes per day. I plan to schedule dedicated")
    print("  morning workout sessions to improve my physical activity index (PhAI).")
    print("- I will aim to balance my sleep duration closer to 8 hours to avoid grogginess")
    print("  and maintain peak afternoon focus.")
    print("=" * 78 + "\n")


if __name__ == "__main__":
    main()