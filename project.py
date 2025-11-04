import pandas as pd
import matplotlib.pyplot as plt

file_path = "Excel sheets/Sheet1.xlsx"
sheet_name = "Data"  # Change if needed
column_name = "Root Cause"  # Change to your column name

# Read Excel
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Extract data from a specific column
column_data = df[column_name]

# Example: print all values
# for value in column_data:
#    print(value)

grouped_df = column_data.value_counts().reset_index()
grouped_df.columns = ["Root Cause", "Count"]
grouped_df.index = grouped_df.index + 1

plt.figure(figsize=(10,6))
plt.bar(grouped_df["Root Cause"], grouped_df["Count"], color="skyblue")
plt.xlabel("Root Cause")
plt.ylabel("Count")
plt.title("Root Cause Frequency")
plt.xticks(rotation=45, ha="right")  # rotate x labels for readability
plt.tight_layout()  # adjust layout to prevent clipping
plt.show()
