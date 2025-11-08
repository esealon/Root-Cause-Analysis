import pandas as pd
import re
import matplotlib.pyplot as plt

file_path = r"Excel sheets/Sheet1.xlsx"
df = pd.read_excel(file_path)

df.columns = df.columns.str.strip()


root_causes_file = r"Excel sheets/Root Causes.xlsx"  # Excel file with the list of root causes
root_causes_df = pd.read_excel(root_causes_file)

# Define key phrases to search for from an Excel sheet
root_causes = root_causes_df["Root Cause"]

# Count occurrences of each phrase
results = []
# Loop through each supplier and each root cause
for supplier in df["Supplier"].unique():
    supplier_data = df[df["Supplier"] == supplier]
    for root_cause in root_causes:
        count = supplier_data["Root Cause"].str.contains(re.escape(root_cause), na=False).sum()
        if count > 0:
            results.append({
                "Supplier": supplier,
                "Root Cause": root_cause,
                "Count": count
            })

# Convert results into a dataframe
supplier_root_cause_counts = pd.DataFrame(results)

# Sort by supplier, then count (descending)
supplier_root_cause_counts = supplier_root_cause_counts.sort_values(
    by=["Supplier", "Count"], ascending=[True, False]
).reset_index(drop=True)

# Start index at 1 instead of 0
supplier_root_cause_counts.index = range(1, len(supplier_root_cause_counts) + 1)

print("\nRoot Cause counts by Supplier:\n")
print(supplier_root_cause_counts)

# Optional: Export results to Excel
output_path = r"Excel sheets/Supplier Root Cause Counts.xlsx"
supplier_root_cause_counts.to_excel(output_path, index=False)
print(f"\nResults saved to {output_path}")

pivot_data = supplier_root_cause_counts.pivot(index="Root Cause", columns="Supplier", values="Count").fillna(0)
pivot_data.plot(kind="bar", figsize=(10, 6))

plt.title("Root Cause Counts by Supplier")
plt.xlabel("Root Cause")
plt.ylabel("Count")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.show()


