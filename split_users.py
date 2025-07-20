import pandas as pd
import os
import re

# === Step 1: This loads the Excel file ===
df = pd.read_excel("users.xlsx")

# This cleans column names
df.columns = df.columns.str.strip()

# === Step 2: This identifies the department column ===
department_column = "Department Name"
df[department_column] = df[department_column].astype(str).str.strip()

# === Step 3: Grouping logic ===
def categorize(dept):
    dept_lower = dept.lower().strip()

    if any(keyword in dept_lower for keyword in ["limpopo"]):
        return "Limpopo Group"
    elif any(keyword in dept_lower for keyword in ["ekurhuleni"]):
        return "City of Ekurhuleni"
    elif any(keyword in dept_lower for keyword in ["corporate"]):
        return "Coporate Services"
    elif any(keyword in dept_lower for keyword in ["coghsta", "settlement", "housing"]):
        return "Human Settlement"
    elif any(keyword in dept_lower for keyword in ["agriculture", "land reform", "rural development"]):
        return "Dep of Agriculture & Rural Development"
    elif any(keyword in dept_lower for keyword in ["community", "safety", "security", "policing"]):
        return "Dep of Community Safety & Security"
    elif any(keyword in dept_lower for keyword in ["sport", "arts", "culture", "recreation"]):
        return "Department of Sports, Arts, Culture and Recreation"
    elif any(keyword in dept_lower for keyword in ["forestry", "fisheries", "environment"]):
        return "Forestry, Fisheries & Environment"
    elif any(keyword in dept_lower for keyword in ["social development"]):
        return "Social Development"
    elif any(keyword in dept_lower for keyword in ["statistics"]):
        return "Statistics SA"
    elif any(keyword in dept_lower for keyword in ["ict", "technology", "information communication"]):
        return "Technology Group"
    elif any(keyword in dept_lower for keyword in ["education", "dhet", "college", "school", "basic education", "higher education"]):
        return "Education Group"
    elif any(keyword in dept_lower for keyword in ["municipality", "metro", "local government"]):
        return "Municipalities"
    elif any(keyword in dept_lower for keyword in ["legislature", "parliament", "provincial"]):
        return "Provincial/Legislative Group"
    elif any(keyword in dept_lower for keyword in ["health", "hospital", "medical"]):
        return "Health Group"
    elif any(keyword in dept_lower for keyword in ["police", "saps"]):
        return "Police Group"
    elif any(keyword in dept_lower for keyword in ["correctional", "correction", "prison", "justice"]):
        return "Department of Correctional Services"
    elif any(keyword in dept_lower for keyword in ["treasury", "finance"]):
        return "National/Provincial Treasury"
    elif any(keyword in dept_lower for keyword in ["tourism", "economic development"]):
        return "Tourism"
    elif any(keyword in dept_lower for keyword in ["water", "sanitation"]):
        return "Water and Sanitation"
    elif "sita" in dept_lower:
        return "SITA"
    elif "nemisa" in dept_lower:
        return "NEMISA"
    elif "post office" in dept_lower:
        return "SA Post Office"
    elif "postbank" in dept_lower:
        return "Postbank"
    elif any(keyword in dept_lower for keyword in ["nyda", "youth development"]):
        return "NYDA"
    elif "home affairs" in dept_lower:
        return "Home Affairs"
    elif any(keyword in dept_lower for keyword in ["transport", "roads", "mobility"]):
        return "Transport"
    elif any(keyword in dept_lower for keyword in ["public works", "infrastructure"]):
        return "Public Works & Infrastructure"
    elif any(keyword in dept_lower for keyword in ["cooperative governance", "cogta"]):
        return "Cooperative Governance"
    elif "national school of government" in dept_lower:
        return "National School of Government"
    elif any(keyword in dept_lower for keyword in ["office of the premier", "premier"]):
        return "Office of the Premier"
    elif "salga" in dept_lower:
        return "SALGA"
    elif "gauteng enterprise propeller" in dept_lower:
        return "Gauteng Enterprise Propeller"
    elif any(keyword in dept_lower for keyword in ["state security", "ssa"]):
        return "State Security Agency"
    elif any(keyword in dept_lower for keyword in ["defence", "sandf", "military"]):
        return "Defence / SANDF"
    elif "government printing works" in dept_lower:
        return "Government Printing Works"
    elif any(keyword in dept_lower for keyword in ["gcis", "government communication"]):
        return "Government Communication & Information System"
    elif any(keyword in dept_lower for keyword in ["sars", "revenue service"]):
        return "South African Revenue Service"
    elif any(keyword in dept_lower for keyword in ["mineral", "petroleum", "petrolium", "resources"]):
        return "Mineral and Petroleum Resources"
    elif dept_lower in [
        "n/a", "na", "none", "non applicable", "not applicable", "nan", "no", "sa", "confirmed",
        "yes", "emfuleni", "dod", "doh", "dot", "daa", "municipality", "n! a", "corporate",
        "office", "sandf", "4-sure", "non", "not working", "", " "]:
        return "N/A Group"
    else:
        unknown_groups.add(dept)
        return "Other/Unclassified"

# === This tracks unknown groups ===
unknown_groups = set()

# === Step 4: This applies categories ===
df["Group"] = df[department_column].apply(categorize)

# === Step 5: This creates output directory ===
output_dir = "department_csvs"
os.makedirs(output_dir, exist_ok=True)

# === Step 6: Define output columns ===
columns_to_keep = ["Contact", "First Name", "Last Name"]

# This validates if columns exist
for col in columns_to_keep:
    if col not in df.columns:
        raise ValueError(f"\u274c Missing expected column: {col}")

# === Step 7: Export grouped CSVs ===
for group_name, group_df in df.groupby("Group"):
    safe_name = re.sub(r'[^\w\-]', '_', group_name.strip())
    output_file = os.path.join(output_dir, f"{safe_name}.csv")
    subset_df = group_df[columns_to_keep].copy()
    subset_df.to_csv(output_file, index=False)
    print(f"\u2705 Saved: {output_file}")

# === Step 8: Save unknown groups ===
if unknown_groups:
    with open("new_groups_to_review.txt", "w") as f:
        for group in sorted(unknown_groups):
            f.write(group + "\n")
    print("\n New or unknown groups are in the 'unclassified' cv group and are also saved to: 'new_groups_to_review.txt' for review.")
else:
    print("\n No unknown groups found.")
