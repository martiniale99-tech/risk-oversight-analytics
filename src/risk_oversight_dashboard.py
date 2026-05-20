import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / "data" / "sample_supervisory_data.csv"
OUTPUT_DIR = BASE_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

df = pd.read_csv(DATA_PATH)

df["planned_start"] = pd.to_datetime(df["planned_start"])
df["planned_end"] = pd.to_datetime(df["planned_end"])

df["duration_days"] = (
    df["planned_end"] - df["planned_start"]
).dt.days

validation_checks = {
    "missing_values": df.isnull().sum().to_dict(),
    "invalid_dates": int(
        (df["planned_end"] < df["planned_start"]).sum()
    ),
    "duplicate_entities": int(
        df["entity_id"].duplicated().sum()
    )
}

kpi_summary = (
    df.groupby(["inspection_type", "status"])
    .size()
    .reset_index(name="count")
)

country_summary = (
    df.groupby("country")["priority_score"]
    .mean()
    .reset_index()
)

country_summary = country_summary.sort_values(
    "priority_score",
    ascending=False
)

excel_output = OUTPUT_DIR / "supervisory_report.xlsx"

with pd.ExcelWriter(
    excel_output,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Clean Data",
        index=False
    )

    kpi_summary.to_excel(
        writer,
        sheet_name="KPI Summary",
        index=False
    )

    country_summary.to_excel(
        writer,
        sheet_name="Country Summary",
        index=False
    )

plt.figure(figsize=(8, 5))

country_summary.plot(
    x="country",
    y="priority_score",
    kind="bar",
    legend=False
)

plt.title("Average Priority Score by Country")
plt.xlabel("Country")
plt.ylabel("Average Priority Score")

plt.tight_layout()

plt.savefig(
    OUTPUT_DIR / "priority_score_dashboard.png"
)

print("Validation checks completed:")
print(validation_checks)

print(f"Excel report saved to: {excel_output}")
