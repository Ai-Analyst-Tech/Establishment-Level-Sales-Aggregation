#AGGREGATING QUANTITIES BY ITEM CODES IN EACH ESTABLISHMENT
from pathlib import Path
import pandas as pd

#Define the correct folder 
base = Path(r"c:\Users\Hp\Desktop")

#Define your input and output files
input_file = base / "20th Oct 25 - 26thOct 25 Weekly Sales Analysis Report.xlsx"  # <-- existing Excel file
output_file = base / "Global Sales Analysis Report.xlsx"  # <-- Excel output file

#Read the data 
data = pd.read_excel(input_file)

#Aggregation
agg_data = (
    data.groupby(["Shop code", "Item code", "Description", "Category"], as_index=False)
    .agg({
        "Quantity": "sum",
        "Unit selling price": "first",
        "Total amount": "sum",
        "Buying Price": "first",
        "Mark Up %": "first"
    })
)

#Save to Excel
agg_data.to_excel(output_file, index=False, engine="openpyxl")
print(f"Aggregation complete! File saved at: {output_file}")