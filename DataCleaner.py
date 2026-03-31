import subprocess
import sys

subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"])

import pandas as pd
import re

df = pd.read_excel("messy_employee_data.xlsx", sheet_name="Employee_Data_RAW")

df.columns = df.columns.str.strip()
df.columns = df.columns.str.lower()
df.columns = df.columns.str.replace(" ", "_")

df = df.rename(columns={
    "employee__id":      "employee_id",
    "email_address":     "email",
    "phone_number":      "phone",
    "salary_$":          "salary",
    "bonus%":            "bonus_percent",
    "country__":         "country"
})

df = df.dropna(how="all")

df = df[~df.apply(lambda row: any(isinstance(v, str) and "---" in v for v in row), axis=1)]

df = df.reset_index(drop=True)

df["employee_id"] = df["employee_id"].astype(str).str.strip().str.replace("EMP-", "")
df = df.drop_duplicates(subset=["employee_id"], keep="first")
df = df.reset_index(drop=True)

df["first_name"] = df["first_name"].astype(str).str.strip().str.title()
df["last_name"]  = df["last_name"].astype(str).str.strip().str.title()

def clean_email(value):
    if pd.isna(value):
        return ""
    value = str(value).strip().lower()
    if "@" not in value:
        return ""
    return value

df["email"] = df["email"].apply(clean_email)

def clean_phone(value):
    if pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 11 and digits[0] == "1":
        digits = digits[1:]
    if len(digits) == 10:
        return f"{digits[0:3]}-{digits[3:6]}-{digits[6:]}"
    return ""

df["phone"] = df["phone"].apply(clean_phone)

df["department"] = df["department"].astype(str).str.strip().str.title()

def clean_salary(value):
    if pd.isna(value):
        return None
    value = str(value).strip().replace("$", "").replace(",", "")
    try:
        return float(value)
    except:
        return None

df["salary"] = df["salary"].apply(clean_salary)

df["hire_date"] = pd.to_datetime(df["hire_date"], dayfirst=False, errors="coerce")
df["hire_date"] = df["hire_date"].dt.strftime("%Y-%m-%d")

def clean_status(value):
    if pd.isna(value):
        return "Unknown"
    value = str(value).strip().title()
    if value == "Active":
        return "Active"
    elif value == "Inactive":
        return "Inactive"
    elif value == "On Leave":
        return "On Leave"
    elif value == "Terminated":
        return "Terminated"
    else:
        return "Unknown"

df["status"] = df["status"].apply(clean_status)

def clean_age(value):
    if pd.isna(value):
        return None
    try:
        age = int(float(str(value)))
        if 18 <= age <= 75:
            return age
        return None
    except:
        return None

df["age"] = df["age"].apply(clean_age)

def clean_gender(value):
    if pd.isna(value):
        return "Not Specified"
    value = str(value).strip().lower()
    if value in ["male", "m"]:
        return "Male"
    elif value in ["female", "f"]:
        return "Female"
    elif value in ["non-binary", "nb"]:
        return "Non-Binary"
    elif value == "other":
        return "Other"
    return "Not Specified"

df["gender"] = df["gender"].apply(clean_gender)

city_fixes = {"la": "Los Angeles", "nyc": "New York"}

def clean_city(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    if value.lower() in city_fixes:
        return city_fixes[value.lower()]
    return value.title()

df["city"] = df["city"].apply(clean_city)

def clean_country(value):
    if pd.isna(value):
        return ""
    value = str(value).strip().lower()
    if value in ["us", "usa", "u.s.a", "united states", "united states of america"]:
        return "USA"
    return value.title()

df["country"] = df["country"].apply(clean_country)

def clean_performance(value):
    if pd.isna(value):
        return None
    value = str(value).strip().replace("%", "")
    try:
        score = float(value)
        if 0 <= score <= 100:
            return score
        return None
    except:
        return None

df["performance_score"] = df["performance_score"].apply(clean_performance)

def clean_bonus(value):
    if pd.isna(value):
        return 0.0
    value = str(value).strip().replace("%", "")
    if value.lower() in ["n/a", "", "none"]:
        return 0.0
    try:
        return float(value)
    except:
        return 0.0

df["bonus_percent"] = df["bonus_percent"].apply(clean_bonus)

df.to_excel("cleaned_employee_data.xlsx", index=False, sheet_name="Cleaned_Data")

print("Done! File saved as cleaned_employee_data.xlsx")
print("Total rows:", len(df))