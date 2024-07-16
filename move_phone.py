import pandas as pd

# Load the Excel files
master_cabang_df = pd.read_excel('master_cabang_demo.xlsx')
head_cabang_df = pd.read_excel('head_cabang_demo.xlsx')

# Clean the NAMA_CABANG column in master_cabang_df, strip and convert to lowercase
master_cabang_df['cleaned_nama_cabang'] = master_cabang_df['NAMA_CABANG'].str.split('/').str[0].str.strip().str.lower()

# Strip and convert the cabang column to lowercase
head_cabang_df['cabang'] = head_cabang_df['cabang'].str.strip().str.lower()

# Create a function to get the correct PIC_NAME and PHONE based on the cabang type
def get_pic_and_phone(row):
    cabang_name = row['cleaned_nama_cabang']
    if pd.isna(cabang_name):
        return pd.Series(['not found', 'not found'])
    if 'capem' in cabang_name:
        pic_name = row['PIC_NAME_CAPEM']
        phone = row['PHONE_CAPEM']
    else:
        pic_name = row['PIC_NAME_CABANG']
        phone = row['PHONE_CABANG']
    return pd.Series([pic_name if pd.notna(pic_name) else 'not found', phone if pd.notna(phone) else 'not found'])

# Merge the dataframes on the cleaned_nama_cabang and cabang columns
merged_df = pd.merge(master_cabang_df, head_cabang_df, left_on='cleaned_nama_cabang', right_on='cabang', how='left')

# Apply the function to get the correct PIC_NAME and PHONE
master_cabang_df[['PIC_NAME', 'PHONE']] = merged_df.apply(get_pic_and_phone, axis=1)

# Convert PHONE to string and add '+62' if it does not already start with '+62' and if it's not 'not found'
master_cabang_df['PHONE'] = master_cabang_df['PHONE'].astype(str).apply(lambda x: f'+62{x}' if x != 'not found' and not x.startswith('+62') else x)

# Drop the temporary cleaned_nama_cabang column
master_cabang_df.drop(columns=['cleaned_nama_cabang'], inplace=True)

# Save the updated master_cabang_df to a new Excel file
output_path = 'updated_master_cabang_new.xlsx'
master_cabang_df.to_excel(output_path, index=False)
