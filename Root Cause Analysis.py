import pandas as pd
import re

file_path = r"Excel sheets/Sheet1.xlsx"
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip()
df["Root Cause"] = df["Root Cause"]

root_causes_file = r"Excel sheets/Root Causes.xlsx"  # Excel file with the list of root causes
root_causes_df = pd.read_excel(root_causes_file)

# Define key phrases to search for from an Excel sheet
root_causes = root_causes_df["Root Cause"]

# Count occurrences of each phrase
results = []
for root_cause in root_causes:
    count = df["Root Cause"].str.contains(re.escape(root_cause), na=False).sum()
    results.append({"Root Cause": root_cause, "Count": count})

root_cause_counts = pd.DataFrame(results).sort_values(by="Count", ascending=False)

root_cause_counts.index = range(1, len(root_cause_counts) + 1)

print("\nPhrase match results:\n")
print(root_cause_counts)


