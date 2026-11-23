import pandas as pd
import numpy as np

# 1. File aur Sheet ke naam set karein
input_file = "Master_Turbine_Data.xlsx"
output_file = "Turbine_Run_Comparison_Result.xlsx"

# Data load karein (1st column Blade Number ko index banaya gaya hai)
df_amp = pd.read_excel(input_file, sheet_name="Amplitude", index_col=0)
df_rpm = pd.read_excel(input_file, sheet_name="RPM", index_col=0)

# Ensure karein ki saare values numeric hain
df_amp = df_amp.apply(pd.to_numeric, errors='coerce')
df_rpm = df_rpm.apply(pd.to_numeric, errors='coerce')

# 2. Run-to-Run Difference Calculate karein (Current Run - Previous Run)
# Example: Run 2 - Run 1, Run 3 - Run 2 ... Run 101 - Run 100
amp_diff = df_amp.diff(axis=1)
rpm_diff = df_rpm.diff(axis=1)

# 3. Status Function: Increase, Decrease ya No Change batane ke liye
def get_trend_status(diff_df):
    status_df = pd.DataFrame(index=diff_df.index, columns=diff_df.columns)
    for col in diff_df.columns:
        conditions = [
            diff_df[col] > 0,
            diff_df[col] < 0,
            diff_df[col] == 0
        ]
        choices = ["Increase", "Decrease", "No Change"]
        status_df[col] = np.select(conditions, choices, default="Baseline (Run 1)")
    return status_df

amp_status = get_trend_status(amp_diff)
rpm_status = get_trend_status(rpm_diff)

# 4. Summary Table banayein (Har Run comparison ka overall nichod)
summary_list = []
runs = df_amp.columns.tolist()

for i in range(1, len(runs)):
    prev_run = runs[i - 1]
    curr_run = runs[i]
    comparison_label = f"{curr_run} vs {prev_run}"
    
    # Blade numbers nikaalein jahan Amplitude badha ya ghata hai
    amp_inc_blades = amp_diff[amp_diff[curr_run] > 0].index.tolist()
    amp_dec_blades = amp_diff[amp_diff[curr_run] < 0].index.tolist()
    
    # Max Amplitude change kaunse blade par aaya
    max_amp_inc_blade = amp_diff[curr_run].idxmax() if not amp_diff[curr_run].isna().all() else None
    max_amp_inc_val = amp_diff[curr_run].max()
    
    summary_list.append({
        "Comparison": comparison_label,
        "Amp_Increased_Blades_Count": len(amp_inc_blades),
        "Amp_Decreased_Blades_Count": len(amp_dec_blades),
        "Max_Amp_Increase_Blade": max_amp_inc_blade,
        "Max_Amp_Delta": round(max_amp_inc_val, 4) if pd.notna(max_amp_inc_val) else 0,
        "RPM_Increased_Blades_Count": (rpm_diff[curr_run] > 0).sum(),
        "RPM_Decreased_Blades_Count": (rpm_diff[curr_run] < 0).sum(),
        "Blades_With_Amp_Increase": ", ".join(map(str, amp_inc_blades))
    })

df_summary = pd.DataFrame(summary_list)

# 5. Sabhi results ko ek nayi Excel file mein alag-alag tabs mein save karein
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_summary.to_excel(writer, sheet_name="Summary_Report", index=False)
    amp_diff.to_excel(writer, sheet_name="Amplitude_Delta_Values")
    amp_status.to_excel(writer, sheet_name="Amplitude_Status")
    rpm_diff.to_excel(writer, sheet_name="RPM_Delta_Values")
    rpm_status.to_excel(writer, sheet_name="RPM_Status")

print("Comparison complete! File saved as:", output_file)