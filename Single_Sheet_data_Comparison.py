import pandas as pd
import numpy as np
from openpyxl.styles import PatternFill

# 1. File aur Sheet ka naam set karein
input_file = "Master_Turbine_Data.xlsx"
sheet_name = "Sheet1"  # Apni master sheet ka exact naam yahan likhein
output_file = "Turbine_SingleSheet_Comparison.xlsx"

# 2. Multi-header (Row 1 = Run No, Row 2 = Amp/RPM) ke saath data load karein
# index_col=0 ka matlab 1st column (Blade Number) index banega
df = pd.read_excel(input_file, sheet_name=sheet_name, header=[0, 1], index_col=0)

# Merged cells ke 'Unnamed' headers ko forward-fill karein taaki har column ko sahi Run No mile
cols = pd.DataFrame(df.columns.tolist())
cols[0] = cols[0].replace(r'^Unnamed:.*', np.nan, regex=True).ffill()
cols[1] = cols[1].astype(str).str.strip()
df.columns = pd.MultiIndex.from_tuples(cols.itertuples(index=False, name=None))

# 3. Amplitude aur RPM ko alag-alag DataFrames mein split karein
# Note: Agar aapki sheet mein 'Amp' likha hai toh 'Amplitude' ki jagah 'Amp' kar dein
df_amp = df.xs('Amplitude', axis=1, level=1).apply(pd.to_numeric, errors='coerce')
df_rpm = df.xs('RPM', axis=1, level=1).apply(pd.to_numeric, errors='coerce')

# 4. Run-to-Run Difference Calculate karein (Current Run - Previous Run)
# (Agar sabhi runs ko sirf Run 1 se compare karna ho toh .diff(axis=1) ki jagah .sub(df_amp.iloc[:, 0], axis=0) use karein)
amp_diff = df_amp.diff(axis=1)
rpm_diff = df_rpm.diff(axis=1)

# 5. Increase / Decrease Status Table
def get_status(diff_df):
    status_df = pd.DataFrame(index=diff_df.index, columns=diff_df.columns)
    for col in diff_df.columns:
        conditions = [diff_df[col] > 0, diff_df[col] < 0, diff_df[col] == 0]
        choices = ["Increase", "Decrease", "No Change"]
        status_df[col] = np.select(conditions, choices, default="Baseline (Run 1)")
    return status_df

amp_status = get_status(amp_diff)
rpm_status = get_status(rpm_diff)

# 6. Summary Report Taiyar Karein (Run-by-Run Overview)
summary_rows = []
runs = df_amp.columns.tolist()

for i in range(1, len(runs)):
    prev_run = runs[i - 1]
    curr_run = runs[i]
    
    amp_inc = amp_diff[amp_diff[curr_run] > 0].index.tolist()
    amp_dec = amp_diff[amp_diff[curr_run] < 0].index.tolist()
    rpm_inc = rpm_diff[rpm_diff[curr_run] > 0].index.tolist()
    rpm_dec = rpm_diff[rpm_diff[curr_run] < 0].index.tolist()
    
    max_amp_blade = amp_diff[curr_run].idxmax() if not amp_diff[curr_run].isna().all() else "-"
    max_amp_val = amp_diff[curr_run].max()
    
    summary_rows.append({
        "Run_Comparison": f"{curr_run} vs {prev_run}",
        "Amp_Inc_Count": len(amp_inc),
        "Amp_Dec_Count": len(amp_dec),
        "Max_Amp_Jump_Blade": max_amp_blade,
        "Max_Amp_Jump_Value": round(max_amp_val, 4) if pd.notna(max_amp_val) else 0,
        "RPM_Inc_Count": len(rpm_inc),
        "RPM_Dec_Count": len(rpm_dec),
        "Blades_Amp_Increased": ", ".join(map(str, amp_inc)),
        "Blades_RPM_Increased": ", ".join(map(str, rpm_inc))
    })

df_summary = pd.DataFrame(summary_rows)

# 7. Output Excel Save Karein + Color Highlighting (Red = Increase, Green = Decrease)
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_summary.to_excel(writer, sheet_name="Summary_Report", index=False)
    amp_diff.to_excel(writer, sheet_name="Amplitude_Delta")
    amp_status.to_excel(writer, sheet_name="Amplitude_Status")
    rpm_diff.to_excel(writer, sheet_name="RPM_Delta")
    rpm_status.to_excel(writer, sheet_name="RPM_Status")
    
    # Status sheets mein color coding lagayein taaki turant dikhe kahan value badhi hai
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # Increase
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Decrease
    
    for sheet in ["Amplitude_Status", "RPM_Status"]:
        ws = writer.sheets[sheet]
        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                if cell.value == "Increase":
                    cell.fill = red_fill
                elif cell.value == "Decrease":
                    cell.fill = green_fill

print(f"Done! Output file saved: {output_file}")