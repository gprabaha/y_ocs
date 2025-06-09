import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import os

# -----------------------------
# Create output folder
# -----------------------------
output_dir = "outputs"
os.makedirs(output_dir, exist_ok=True)

# -----------------------------
# Load and filter data
# -----------------------------
df = pd.read_excel("01_AlumData.xlsx")

# Filter BBS rows
bbs_df = df[df['GPR_Program'].astype(str).str.startswith("BBS:")].copy()

# Relevant columns
columns_to_export = [
    'GPR_Acd Yr',
    'GPR_Division',
    'GPR_Program',
    'GPR_Most Recent Known Employer',
    'GPR_Most Recent Known Employment Position',
    'AAUDE_Major_Industry',
    'AAUDE_Major_Position'
]

# Check for missing columns
available_cols = [col for col in columns_to_export if col in bbs_df.columns]
missing_cols = set(columns_to_export) - set(available_cols)
if missing_cols:
    print(f"⚠️ Missing columns: {missing_cols}")

# Clean and standardize missing values
export_df = bbs_df[available_cols].copy()
export_df.replace("", pd.NA, inplace=True)
export_df.fillna("Missing", inplace=True)

# Save cleaned Excel
excel_path = os.path.join(output_dir, "bbs_current_positions.xlsx")
export_df.to_excel(excel_path, index=False)
print(f"✅ Exported cleaned data to: {excel_path}")

# -----------------------------
# Visualization Functions
# -----------------------------

def save_pie_chart(series, title, filename, min_pct=2):
    series = series[series != "Missing"]
    total = len(series)
    if total == 0:
        print(f"⚠️ {title}: No non-missing data to plot.")
        return
    counts = series.value_counts()
    counts = counts.sort_values(ascending=False)

    major = counts[counts / total >= min_pct / 100]
    other = counts[counts / total < min_pct / 100]
    if not other.empty:
        major['Other'] = other.sum()

    plt.figure(figsize=(6, 6))
    major.plot(kind='pie', labels=major.index, autopct='%1.1f%%', startangle=140)
    plt.title(title)
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, filename))
    plt.close()

def save_bar_chart(series, title, filename, min_count=5):
    original_len = len(series)
    series = series[series != "Missing"]
    counts = series.value_counts()
    missing_pct = 100 * (original_len - len(series)) / original_len
    if missing_pct > 0:
        print(f"⚠️ {title}: Dropped {missing_pct:.1f}% missing entries.")

    counts = counts[counts >= min_count]
    if counts.empty:
        print(f"⚠️ {title}: No values with count >= {min_count}. Skipping plot.")
        return

    plt.figure(figsize=(10, max(4, 0.3 * len(counts))))
    counts.sort_values().plot(kind='barh')
    plt.title(title)
    plt.xlabel("Number of Individuals")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, filename))
    plt.close()

# -----------------------------
# Summary Visualizations
# -----------------------------

save_pie_chart(
    export_df['AAUDE_Major_Industry'],
    title="Distribution of AAUDE Major Industry",
    filename="aauude_major_industry_pie.png"
)

save_bar_chart(
    export_df['GPR_Most Recent Known Employer'],
    title="Distribution of Most Recent Known Employers",
    filename="all_employers_bar.png"
)

save_bar_chart(
    export_df['GPR_Most Recent Known Employment Position'],
    title="Distribution of Most Recent Known Positions",
    filename="all_positions_bar.png"
)

save_bar_chart(
    export_df['AAUDE_Major_Position'],
    title="Distribution of AAUDE Major Position",
    filename="aauude_major_position_bar.png"
)

# -----------------------------
# Sankey Diagram (Scrollable HTML)
# -----------------------------

# Remove any rows with "Missing"
df_sankey = export_df[
    ['GPR_Program', 'AAUDE_Major_Industry', 'GPR_Most Recent Known Employer', 'GPR_Most Recent Known Employment Position']
].copy()

initial_count = len(df_sankey)
df_sankey = df_sankey[~df_sankey.eq("Missing").any(axis=1)]
excluded_pct = 100 * (initial_count - len(df_sankey)) / initial_count
if excluded_pct > 0:
    print(f"⚠️ Sankey: Dropped {excluded_pct:.1f}% of rows with missing data.")

# Collapse rare categories into 'Other'
def collapse_rare(df, col, min_count=5):
    counts = df[col].value_counts()
    rare = counts[counts < min_count].index
    df[col] = df[col].apply(lambda x: "Other" if x in rare else x)
    return df

for col in df_sankey.columns:
    df_sankey = collapse_rare(df_sankey, col, min_count=5)

# Rebuild node list
levels = list(df_sankey.columns)
label_indices = {}
all_labels = []
idx = 0

for level in levels:
    for val in df_sankey[level].unique():
        key = (level, val)
        if key not in label_indices:
            label_indices[key] = idx
            all_labels.append(val)
            idx += 1

# Build links
source = []
target = []
value = []

for i in range(len(levels) - 1):
    grouped = df_sankey.groupby([levels[i], levels[i+1]]).size().reset_index(name='count')
    for _, row in grouped.iterrows():
        src = label_indices[(levels[i], row[levels[i]])]
        tgt = label_indices[(levels[i+1], row[levels[i+1]])]
        source.append(src)
        target.append(tgt)
        value.append(row['count'])

# Create scrollable Sankey layout
fig = go.Figure(data=[go.Sankey(
    arrangement="snap",
    node=dict(
        pad=15,
        thickness=15,
        label=all_labels,
        color="lightblue"
    ),
    link=dict(
        source=source,
        target=target,
        value=value
    )
)])

fig.update_layout(
    title_text="BBS Alumni Flow: Program → Industry → Employer → Position",
    font_size=10,
    autosize=False,
    width=1800,
    height=max(800, 30 * len(all_labels)),
)

fig.write_html(os.path.join(output_dir, "bbs_flow_sankey.html"))
print("✅ Sankey diagram saved as scrollable HTML.")
print("📊 All outputs saved in 'outputs/'")
