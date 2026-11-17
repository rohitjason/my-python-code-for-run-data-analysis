import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ==============================================================================
# 1. SETTINGS & FILTERS (Yahan apni requirements set karein)
# ==============================================================================
INPUT_FILE = "Run01_Data.xlsx"
SHEET_NAME = 0

# Column Names (Excel ke exact column names likhein)
COL_BLADE = "Blade_No"
COL_RPM = "RPM"
COL_AMP = "Amplitude"

# --- NEW: RPM RANGE FILTER ---
FILTER_BY_RPM = True  # False karein agar poora RPM range chahiye
MIN_RPM = 8000  # Minimum RPM limit
MAX_RPM = 10500  # Maximum RPM limit

# --- NEW: ENGINE ORDER (EO) FILTER ---
FILTER_BY_EO = True  # False karein agar sabhi EOs ka combined data chahiye
TARGET_EO = 3  # Kaunsa Engine Order check karna hai (e.g., 2, 3, 4, 5...)
EO_TOLERANCE = 0.15  # ±0.15 tolerance (e.g., 2.85 se 3.15 tak ko 3EO manega)

# Agar Excel mein 'EO' column pehle se hai toh uska naam likhein, warna None rakhein:
COL_EO = "EO"  # e.g., "EO" ya "Order" (agar nahi hai toh None likhein)
# Agar 'EO' column nahi hai par 'Frequency' (Hz) column hai, toh uska naam likhein:
COL_FREQ = None  # e.g., "Frequency_Hz" (agar nahi hai toh None rakhein)

OUTPUT_EXCEL = f"Run_Report_RPM_{MIN_RPM}_{MAX_RPM}_EO{TARGET_EO}.xlsx"
OUTPUT_IMAGE = f"Run_Charts_RPM_{MIN_RPM}_{MAX_RPM}_EO{TARGET_EO}.png"

# ==============================================================================
# 2. DATA LOAD & CLEANING
# ==============================================================================
print(f"Loading {INPUT_FILE}...")
df = (
    pd.read_csv(INPUT_FILE)
    if INPUT_FILE.endswith(".csv")
    else pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
)

df[COL_BLADE] = pd.to_numeric(df[COL_BLADE], errors="coerce")
df[COL_RPM] = pd.to_numeric(df[COL_RPM], errors="coerce")
df[COL_AMP] = pd.to_numeric(df[COL_AMP], errors="coerce")
df = df.dropna(subset=[COL_BLADE, COL_RPM, COL_AMP])
df[COL_BLADE] = df[COL_BLADE].astype(int)

# ==============================================================================
# 3. APPLY RPM RANGE FILTER
# ==============================================================================
if FILTER_BY_RPM:
  df = df[(df[COL_RPM] >= MIN_RPM) & (df[COL_RPM] <= MAX_RPM)]
  print(f"Filtered RPM [{MIN_RPM} - {MAX_RPM}]: {len(df)} rows remaining.")

# ==============================================================================
# 4. APPLY ENGINE ORDER (EO) FILTER
# ==============================================================================
if FILTER_BY_EO:
  if COL_EO and COL_EO in df.columns:
    # Direct EO column se filter karna
    df[COL_EO] = pd.to_numeric(df[COL_EO], errors="coerce")
    df = df[np.abs(df[COL_EO] - TARGET_EO) <= EO_TOLERANCE]
    print(f"Filtered by {TARGET_EO}EO (from '{COL_EO}' column): {len(df)} rows.")

  elif COL_FREQ and COL_FREQ in df.columns:
    # Frequency aur RPM se EO calculate karke filter karna: EO = (Freq * 60) / RPM
    df[COL_FREQ] = pd.to_numeric(df[COL_FREQ], errors="coerce")
    df["Calculated_EO"] = (df[COL_FREQ] * 60.0) / df[COL_RPM]
    df = df[np.abs(df["Calculated_EO"] - TARGET_EO) <= EO_TOLERANCE]
    print(
        f"Filtered by {TARGET_EO}EO (calculated from Frequency): {len(df)} rows."
    )
  else:
    print(
        "WARNING: Excel mein na 'EO' column mila na 'Frequency' column. EO filter"
        " skip kiya gaya."
    )

if df.empty:
  raise ValueError(
      "Selected RPM Range ya Engine Order ke liye koi data nahi mila! Limits"
      " check karein."
  )

# ==============================================================================
# 5. BLADE-WISE SUMMARY (BLADE 1 TO 47)
# ==============================================================================
idx_max = df.groupby(COL_BLADE)[COL_AMP].idxmax()
peak_df = df.loc[idx_max, [COL_BLADE, COL_AMP, COL_RPM]].copy()
peak_df.rename(
    columns={COL_AMP: "Max_Amplitude", COL_RPM: "Peak_RPM"}, inplace=True
)

stats_df = (
    df.groupby(COL_BLADE)[COL_AMP]
    .agg(
        Mean_Amplitude="mean",
        RMS_Amplitude=lambda x: np.sqrt(np.mean(x**2)),
    )
    .reset_index()
)

summary_df = pd.merge(peak_df, stats_df, on=COL_BLADE).sort_values(
    by=COL_BLADE
)

# Anomaly Threshold (Mean + 2*SD)
mean_val = summary_df["Max_Amplitude"].mean()
std_val = summary_df["Max_Amplitude"].std()
threshold = mean_val + (2 * std_val)
summary_df["Status"] = np.where(
    summary_df["Max_Amplitude"] > threshold, "HIGH ALERT", "Normal"
)

# Save Excel
with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
  summary_df.to_excel(writer, sheet_name="Filtered_Blade_Summary", index=False)
  summary_df.sort_values(by="Max_Amplitude", ascending=False).head(10).to_excel(
      writer, sheet_name="Top_10_Blades", index=False
  )

print(f"Report saved: {OUTPUT_EXCEL}")

# ==============================================================================
# 6. PLOTTING (BAR CHART + CAMPBELL / EO LINE)
# ==============================================================================
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 9))
title_eo = f" | Engine Order: {TARGET_EO}EO" if FILTER_BY_EO else ""
title_rpm = f" | RPM: {MIN_RPM}–{MAX_RPM}" if FILTER_BY_RPM else ""
fig.suptitle(
    f"Fan 2R (47 Blades) Analysis{title_rpm}{title_eo}",
    fontsize=14,
    fontweight="bold",
)

# Top Plot: 47 Blades Bar Chart
colors = [
    "red" if x > threshold else "#2b5c8f" for x in summary_df["Max_Amplitude"]
]
ax1.bar(
    summary_df[COL_BLADE],
    summary_df["Max_Amplitude"],
    color=colors,
    edgecolor="black",
)
ax1.axhline(
    mean_val, color="green", linestyle="--", label=f"Mean ({mean_val:.2f})"
)
ax1.axhline(
    threshold,
    color="red",
    linestyle=":",
    label=f"Alert Threshold ({threshold:.2f})",
)
ax1.set_title("Max Amplitude per Blade in Selected RPM & EO Window")
ax1.set_xlabel("Blade Number (1 to 47)")
ax1.set_ylabel("Amplitude")
ax1.set_xticks(range(1, 48))
ax1.tick_params(axis="x", labelsize=8)
ax1.grid(axis="y", linestyle="--", alpha=0.5)
ax1.legend()

# Bottom Plot: Amplitude vs RPM (Top 3 Blades + Theoretical EO Line if Freq available)
top3 = (
    summary_df.sort_values(by="Max_Amplitude", ascending=False)
    .head(3)[COL_BLADE]
    .tolist()
)
for b in top3:
  b_sub = df[df[COL_BLADE] == b].sort_values(by=COL_RPM)
  ax2.plot(
      b_sub[COL_RPM],
      b_sub[COL_AMP],
      marker="o",
      markersize=4,
      label=f"Blade #{b}",
  )

ax2.set_title(
    f"Amplitude vs. RPM Response (Top 3 Blades in {MIN_RPM}-{MAX_RPM} RPM)"
)
ax2.set_xlabel("RPM")
ax2.set_ylabel("Amplitude")
ax2.grid(True, linestyle="--", alpha=0.5)
ax2.legend()

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig(OUTPUT_IMAGE, dpi=300)
print(f"Chart saved: {OUTPUT_IMAGE}")
plt.show()