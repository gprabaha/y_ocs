import pandas as pd
import matplotlib.pyplot as plt
import os
import re

# -------------------------
# CONFIGURATION
# -------------------------

# File path
input_file = "01_AlumData.xlsx"

# Columns to summarize
columns_of_interest = [
    'GPR_Cur Emplymt Status',
    'GPR_Cur Sector',
    'GPR_Cur Emplymt Type',
    'GPR_Most Recent Known Employer',
    'GPR_Most Recent Known Employment Position',
    'GPR_Lst Known Fac Status'
]

# -------------------------
# FUNCTIONS
# -------------------------

def safe_filename(name):
    return re.sub(r'[^\w\-_\. ]', '', str(name))[:40]

def analyze_group(df, output_folder, group_label):
    os.makedirs(output_folder, exist_ok=True)

    # --- 1. Employment Type by Sector ---
    sectors = df['GPR_Cur Sector'].dropna().unique()

    for sector in sectors:
        sub_df = df[df['GPR_Cur Sector'] == sector]
        emp_type_counts = sub_df['GPR_Cur Emplymt Type'].value_counts(dropna=False)

        if emp_type_counts.empty:
            continue

        ax = emp_type_counts.plot(kind='bar', figsize=(10, 6))
        plt.title(f"{group_label} - Employment Types in Sector: {sector}", fontsize=12)
        plt.ylabel("Number of Individuals")
        plt.xticks(rotation=45, ha='right', fontsize=8)
        plt.tight_layout()
        clean_sector = safe_filename(sector)
        plt.savefig(os.path.join(output_folder, f"employment_type_in_sector__{clean_sector}.png"))
        plt.close()

    # --- 2. Summary bar plots for other columns (Top 10) ---
    for col in columns_of_interest:
        if col in ['GPR_Cur Sector', 'GPR_Cur Emplymt Type']:
            continue

        vc = df[col].value_counts(dropna=False).head(10)

        if vc.empty:
            continue

        ax = vc.plot(kind='barh', figsize=(10, 6))
        plt.title(f"{group_label} - Top 10 Values for: {col}", fontsize=12)
        plt.xlabel("Number of Individuals")
        plt.yticks(fontsize=8)
        plt.tight_layout()
        fname = f"summary__{safe_filename(col)}.png"
        plt.savefig(os.path.join(output_folder, fname))
        plt.close()


# -------------------------
# MAIN SCRIPT
# -------------------------

# Load data
df = pd.read_excel(input_file)

# Filter for BBS rows
bbs_df = df[df['GPR_Program'].astype(str).str.startswith('BBS:')].copy()

# ---- 1. Overall BBS Analysis ----
analyze_group(bbs_df, "plots/overall", group_label="BBS overall")

# ---- 2. Per BBS Subtype Analysis ----
bbs_subtypes = bbs_df['GPR_Program'].dropna().unique()

for subtype in bbs_subtypes:
    sub_df = bbs_df[bbs_df['GPR_Program'] == subtype]
    folder_name = f"plots/{safe_filename(subtype)}"
    analyze_group(sub_df, folder_name, group_label=subtype)
