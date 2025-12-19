import pandas as pd
import matplotlib.pyplot as plt

# -INPUT DATA AS A CSV FILE-
data = pd.read_csv("production_data3.csv")

# -DECLARING STANDARD VALUES-
STANDARDS = {
    "Cycle_Time_sec": 12,
    "Idle_Time_sec": 0,
    "Defect_Count": 0,
    "Distance_Moved_m": 5,
    "Motion_Time_sec": 6
            }

WASTE_MAP = {
    "Cycle_Time_sec": "WAITING",
    "Idle_Time_sec": "WAITING",
    "Defect_Count": "DEFECTS",
    "Distance_Moved_m": "MOTION",
    "Motion_Time_sec": "MOTION"
            }

# -WASTE ANALYSIS FUNCTION-
def analyze_row(row):
    waste_hits = []
    for col, limit in STANDARDS.items():
        value = row[col]
        if pd.isna(value):
            continue
        if value > limit:
            waste_hits.append(WASTE_MAP[col])
    count = len(waste_hits)
    if count > 4:
        level = "CRITICAL"
    elif count > 2:
        level = "WASTE"
    else:
        level = "OPTIMAL"
    if waste_hits:
        return f"{level} | {', '.join(sorted(set(waste_hits)))}"
    else:
        return level

# -ANALYSISING THE INPUT DATA-
data["PROCESS-STATUS"] = data.apply(analyze_row, axis=1)
print("\n - ANALYZED DATA - \n")
print(data)

# -CATEGORIZING THE WASTE-
waste_series = (
    data["PROCESS-STATUS"]
    .str.split("|")
    .str[-1]
    .str.strip()
)

waste_series = waste_series[waste_series != "OPTIMAL"]
waste_categories = waste_series.str.split(", ").explode()

# -VISUAL REPRESENTATION USING PIE CHART-
if not waste_categories.empty:
    
    WASTE_COLORS = {
        "WAITING": "#FFD966",    # Yellow – delay / idle
        "MOTION": "#9DC3E6",     # Blue – movement
        "DEFECTS": "#F4B084"     # Orange/Red – quality issue
    }

    waste_counts = waste_categories.value_counts()

    # SAFE color mapping (prevents KeyError)
    colors = [WASTE_COLORS.get(w, "#D9D9D9") for w in waste_counts.index]

    plt.figure()
    plt.pie(
        waste_counts,
        labels=waste_counts.index,
        colors=colors,
        autopct="%1.1f%%",
        startangle=90,
        wedgeprops={"edgecolor": "black"}  # fixed
    )
    plt.title(" - WASTE DISTRIBUTIONS - ")
    plt.show()

else:
    print("\n - NO WASTE DETECTED - ")

# -LEAN BASIC PROBLEMS AND THEIR SOLUTIONS-
LEAN_SOLUTIONS = {
    "WAITING": {
        "Immediate-Action": "Visual job sequencing; Temporary workload balancing",
        "RootCause-Tools": "Time study; Bottleneck analysis",
        "LongTerm-Improvements": "Line balancing; Pull-based production"
    },
    "MOTION": {
        "Immediate-Action": "Reposition tools; Reduce unnecessary movement",
        "RootCause-Tools": "Motion study; Ergonomic assessment",
        "LongTerm-Improvements": "Layout redesign; 5S implementation"
    },
    "DEFECTS": {
        "Immediate-Action": "Segregate defective parts; Temporary inspection",
        "RootCause-Tools": "5-Why analysis; Fishbone diagram",
        "LongTerm-Improvements": "Poka-Yoke; SPC implementation"
    }
}

# - GIVING SOLUTIONS TO UNQIUE PROBLEM-
unique_wastes = sorted(set(waste_categories))
solution_table = pd.DataFrame([
    {
        "Waste_Type": w,
        "Immediate-Action": LEAN_SOLUTIONS[w]["Immediate-Action"],
        "RootCause-Tools": LEAN_SOLUTIONS[w]["RootCause-Tools"],
        "LongTerm-Improvements": LEAN_SOLUTIONS[w]["LongTerm-Improvements"]
    }
    for w in unique_wastes
    if w in LEAN_SOLUTIONS
])

print("\n  -LEAN SOLUTIONS- \n")
print(solution_table)

# - WASTE RANKING-
SEVERITY_SCORE = {
    "OPTIMAL": 0,
    "WASTE": 1,
    "CRITICAL": 2
}
TIME_COLUMNS = [col for col in STANDARDS if "Time" in col]

time_deviation = {}
for col in TIME_COLUMNS:
    deviation = (data[col] - STANDARDS[col]).clip(lower=0)
    time_deviation[col] = deviation.sum()

priority_rows = []

if not waste_categories.empty:
    frequency_table = waste_categories.value_counts()

    for waste, freq in frequency_table.items():

        time_impact = 0
        for col, mapped_waste in WASTE_MAP.items():
            if mapped_waste == waste and col in time_deviation:
                time_impact += time_deviation[col]
        related_rows = data["PROCESS-STATUS"].str.contains(waste, na=False)
        urgency = "OPTIMAL"
        if data.loc[related_rows, "PROCESS-STATUS"].str.contains("CRITICAL").any():
            urgency = "CRITICAL"
        elif data.loc[related_rows, "PROCESS-STATUS"].str.contains("WASTE").any():
            urgency = "WASTE"
        priority_score = freq + time_impact + SEVERITY_SCORE[urgency]
        priority_rows.append({
            "Waste_Type": waste,
            "Frequency": freq,
            "Time_Impact": round(time_impact, 2),
            "Urgency": urgency,
            "Priority_Score": round(priority_score, 2)
        })

# - RANK TABLE FOR WASTES -
priority_table = (
    pd.DataFrame(priority_rows)
    .sort_values(by="Priority_Score", ascending=False)
    .reset_index(drop=True)
)
print("\n--- WASTE PRIORITY RANKING (DECISION VIEW) ---\n")
print(priority_table)

def enhance_excel_sheet(workbook, worksheet, df, highlight_column=None):
    header_fmt = workbook.add_format({
        "bold": True,
        "border": 1,
        "align": "center",
        "valign": "middle",
        "bg_color": "#BDD7EE",
        "text_wrap": True
    })

    normal_fmt = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "middle"
    })

    wrap_fmt = workbook.add_format({
        "border": 1,
        "text_wrap": True,
        "valign": "top"
    })

    critical_fmt = workbook.add_format({
        "bg_color": "#F8CBAD",
        "border": 1,
        "align": "center"
    })

    waste_fmt = workbook.add_format({
        "bg_color": "#FFE699",
        "border": 1,
        "align": "center"
    })

    ok_fmt = workbook.add_format({
        "bg_color": "#C6E0B4",
        "border": 1,
        "align": "center"
    })

    for col_idx, col_name in enumerate(df.columns):
        max_len = max(
            df[col_name].astype(str).map(len).max(),
            len(col_name)
        )
        worksheet.set_column(col_idx, col_idx, min(max_len + 3, 45))
        worksheet.write(0, col_idx, col_name, header_fmt)

        for row_idx in range(len(df)):
            value = df.iloc[row_idx, col_idx]
            if pd.isna(value):
                worksheet.write_blank(row_idx + 1, col_idx, None, normal_fmt)
                continue
            if highlight_column == col_name:
                if "CRITICAL" in str(value):
                    fmt = critical_fmt
                elif "WASTE" in str(value):
                    fmt = waste_fmt
                elif "OPTIMAL" in str(value):
                    fmt = ok_fmt
                else:
                    fmt = normal_fmt
            else:
                fmt = wrap_fmt if isinstance(value, str) and len(value) > 25 else normal_fmt

            worksheet.write(row_idx + 1, col_idx, value, fmt)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

# - EXCEL REPORT WITH CLEAN FORMAT-
report_file = "Lean_Waste_Analysis_Report.xlsx"

with pd.ExcelWriter(report_file, engine="xlsxwriter") as writer:
    workbook = writer.book

    # Sheet 1: Analyzed Data
    data.to_excel(writer, sheet_name="Analyzed_Data", index=False)
    enhance_excel_sheet(
        workbook,
        writer.sheets["Analyzed_Data"],
        data,
        highlight_column="PROCESS-STATUS"
    )

    # Sheet 2: Waste Summary
    waste_summary = waste_categories.value_counts().reset_index()
    waste_summary.columns = ["Waste_Type", "Frequency"]

    waste_summary.to_excel(writer, sheet_name="Waste_Summary", index=False)
    enhance_excel_sheet(
        workbook,
        writer.sheets["Waste_Summary"],
        waste_summary
    )

    # Sheet 3: Waste Priority
    priority_table.to_excel(writer, sheet_name="Waste_Priority", index=False)
    enhance_excel_sheet(
        workbook,
        writer.sheets["Waste_Priority"],
        priority_table,
        highlight_column="Urgency"
    )

    # Sheet 4: Lean Solutions
    solution_table.to_excel(writer, sheet_name="Lean_Solutions", index=False)
    enhance_excel_sheet(
        workbook,
        writer.sheets["Lean_Solutions"],
        solution_table
    )

    # Sheet 5: Charts
    chart_sheet = workbook.add_worksheet("Charts")

    pie_chart = workbook.add_chart({"type": "pie"})
    pie_chart.add_series({
        "categories": "=Waste_Summary!$A$2:$A${}".format(len(waste_summary) + 1),
        "values": "=Waste_Summary!$B$2:$B${}".format(len(waste_summary) + 1),
        "data_labels": {"percentage": True},
    })

    pie_chart.set_title({"name": "Categorized Waste Distribution"})
    chart_sheet.insert_chart("B2", pie_chart, {"x_scale": 1.5, "y_scale": 1.5})

print(f"\nExcel report generated successfully: {report_file}")
