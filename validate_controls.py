import pandas as pd

INPUT_FILE = "internal_control_raw.xlsx"
OUTPUT_FILE = "internal_control_validated.xlsx"

def parse_range(rng: str):
    # "10-50" -> (10, 50)
    try:
        low, high = rng.split("-")
        return int(low), int(high)
    except Exception:
        return None, None

df = pd.read_excel(INPUT_FILE)

# Create error columns
df["ERR_NotSubmitted"] = df["Submitted"].astype(str).str.strip().str.lower().ne("yes")
df["ERR_MissingValue"] = df["Value"].isna()

# Range check
low_high = df["Expected_Range"].astype(str).apply(parse_range)
df["Expected_Min"] = [x[0] for x in low_high]
df["Expected_Max"] = [x[1] for x in low_high]

df["ERR_RangeInvalid"] = df["Expected_Min"].isna() | df["Expected_Max"].isna()
df["ERR_OutOfRange"] = (~df["ERR_RangeInvalid"]) & (
    (df["Value"] < df["Expected_Min"]) | (df["Value"] > df["Expected_Max"])
)

# Total error flag
error_cols = ["ERR_NotSubmitted", "ERR_MissingValue", "ERR_RangeInvalid", "ERR_OutOfRange"]
df["Has_Error"] = df[error_cols].any(axis=1)

# KPI summary
total = len(df)
submitted_rate = (df["Submitted"].astype(str).str.strip().str.lower().eq("yes").sum() / total) * 100
error_rate = (df["Has_Error"].sum() / total) * 100

print("=== KPI SUMMARY ===")
print(f"Total rows: {total}")
print(f"Submission rate: {submitted_rate:.1f}%")
print(f"Error rate: {error_rate:.1f}%")
print("\nTop error counts:")
print(df[error_cols].sum().sort_values(ascending=False))

print("\n=== COUNTRY SUMMARY ===")
country_summary = df.groupby("Country").agg(
    Total=("Country", "count"),
    Submitted_Yes=("Submitted", lambda x: (x.astype(str).str.lower().eq("yes")).sum()),
    Errors=("Has_Error", "sum")
)
country_summary["SubmissionRate_%"] = (country_summary["Submitted_Yes"] / country_summary["Total"] * 100).round(1)
country_summary["ErrorRate_%"] = (country_summary["Errors"] / country_summary["Total"] * 100).round(1)
print(country_summary.sort_values("ErrorRate_%", ascending=False))

# Save output
df.to_excel(OUTPUT_FILE, index=False)
print(f"\nSaved validated file: {OUTPUT_FILE}")
