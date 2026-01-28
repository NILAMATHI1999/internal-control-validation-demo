import pandas as pd
from datetime import datetime, timedelta
import random

countries = ["Germany", "India", "USA", "France", "Brazil"]
control_ids = ["IC101", "IC102", "IC103", "IC104", "IC105"]

rows = []

for i in range(50):
    country = random.choice(countries)
    control = random.choice(control_ids)
    submitted = random.choice(["Yes", "Yes", "Yes", "No"])  # mostly yes
    value = random.randint(0, 100)
    expected_min = random.randint(10, 30)
    expected_max = expected_min + random.randint(20, 40)
    date = datetime.now() - timedelta(days=random.randint(0, 10))

    rows.append([
        country,
        control,
        submitted,
        value,
        f"{expected_min}-{expected_max}",
        date.strftime("%d-%m-%Y")
    ])

df = pd.DataFrame(rows, columns=[
    "Country",
    "Control_ID",
    "Submitted",
    "Value",
    "Expected_Range",
    "Date"
])

df.to_excel("internal_control_raw.xlsx", index=False)

print("Excel file created: internal_control_raw.xlsx")

