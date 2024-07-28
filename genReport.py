import pandas as pd

# Load the Excel file
file_path = 'HISTORY.xlsx'
df = pd.read_excel(file_path)

# Convert the 'TANGGAL INPUT' column to datetime
df['TANGGAL INPUT'] = pd.to_datetime(df['TANGGAL INPUT'], format='%d/%m/%Y %H:%M:%S')

# Filter data for today only
today_date = pd.to_datetime('today').normalize()
df_today = df[df['TANGGAL INPUT'].dt.normalize() == today_date]

# Convert the 'KETERANGAN' and 'STATUS' columns to strings after handling NaNs
df_today['KETERANGAN'] = df_today['KETERANGAN'].fillna('').astype(str)
df_today['PROGRES_PERBAIKAN_ATM'] = df_today['PROGRES_PERBAIKAN_ATM'].fillna('').astype(str)

# Define the sections
sections = {
    "*Berikut List ATM down hingga sore hari ini:*": df_today[
        (df_today['TIPE_PERMASALAHAN'] == 'Problem Down') & (df_today['PROGRES_PERBAIKAN_ATM'].isnull())
    ],
    "*Berikut List ATM yang belum dapat respond dari pihak pengelola:*": df_today[
        (df_today['KETERANGAN'] == '') & (df_today['PROGRES_PERBAIKAN_ATM'] == '')
    ],
    "*Berikut List ATM yang sedang dalam proses pengerjaan:*": df_today[
        (df_today['KETERANGAN'].str.contains('in progress', case=False)) & (df_today['PROGRES_PERBAIKAN_ATM'] == '')
    ],
    "*Berikut List ATM yang mendapatkan error hingga sore ini namun setelah dilakukan pengecekan oleh pihak cabang, ATM berjalan normal:*": df_today[
        (df_today['KETERANGAN'] != '') & (~df_today['KETERANGAN'].str.contains('in progress', case=False)) & (df_today['PROGRES_PERBAIKAN_ATM'] == '')
    ]
}

# Construct the report
report = "Selamat sore, izin untuk report status ATM hingga sore ini\n\n"

for section_title, data in sections.items():
    report += section_title + "\n\n"
    if not data.empty:
        for idx, (_, row) in enumerate(data.iterrows(), start=1):
            report += f"{idx}. {row['NAMA_ATM']} {row['PERMASALAHAN']}\n"
    else:
        report += "Tidak ada data.\n"
    report += "\n"

# Save the report to a text file
report_path = 'ATM_Report.txt'
with open(report_path, 'w') as file:
    file.write(report)

print(f"Report saved to {report_path}")
