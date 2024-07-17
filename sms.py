import pandas as pd
import re
from datetime import datetime
import urllib.parse
from collections import defaultdict

# Function to read and process the text file
def process_text_file(text_file_path, atm_info, exceptions):
    problems = []
    not_found = []
    above_ten_percent = []
    atm_name_patterns = [
        r'ATM\s[\w\s\d\(\)-]+',
        r'CRM\s[\w\s\d\(\)-]+'
    ]

    with open(text_file_path, 'r') as file:
        lines = file.readlines()
        monitoring_npm_section = False
        saldo_pagu_section = False
        atm_problem_section = False
        error_type = ""

        for line in lines:
            if 'Problem Hardware' in line:
                error_type = 'Problem Hardware'
            elif 'Problem Down' in line:
                error_type = 'Problem Down'
            elif 'Problem Supply Out' in line:
                error_type = 'Problem Supply Out'
            elif 'ATM Warning' in line:
                error_type = 'ATM Warning'
            elif 'Report Persentase Saldo di Bawah Pagu ATM BPD Bali' in line:
                error_type = 'Saldo di Bawah Pagu'
                saldo_pagu_section = True
                monitoring_npm_section = False
                atm_problem_section = False
                continue

            if 'monitoring_npm:' in line:
                error_type = 'NPM Problem'
                monitoring_npm_section = True
                saldo_pagu_section = False
                atm_problem_section = False
                continue
            
            if 'Report Problem ATM BPD Bali' in line:
                atm_problem_section = True
                monitoring_npm_section = False
                saldo_pagu_section = False
                continue

            # Process lines differently based on the section
            if monitoring_npm_section:
                for pattern in atm_name_patterns:
                    match = re.search(pattern, line)
                    if match:
                        atm_name = match.group().strip()
                        # Match atm_name with NAMA_ATM from atm_info DataFrame, ignoring case and trailing spaces
                        atm_match = atm_info[atm_info['NAMA_ATM'].str.strip().str.casefold() == atm_name.casefold()]
                        if not atm_match.empty:
                            id_atm = atm_match.iloc[0]['ID_ATM']
                            id_atm_str = f"{int(id_atm):08d}"  # Ensure ATM ID is 8 digits
                            # Check if id_atm is in exceptions
                            if id_atm not in exceptions:
                                problems.append({"ID_ATM": id_atm_str, "NAMA_ATM": atm_name, "PROBLEM": f"error dengan keterangan : ID ATM {id_atm_str} Down Node - No further details", "TYPE": error_type})
                        else:
                            not_found.append({"ATM_NAME": atm_name, "PROBLEM": "Down Node - No further details", "TYPE": error_type})
                        break
            elif saldo_pagu_section:
                match = re.match(r'\s*\d+\.\s*(\d+)\s*\|\s*([^|]+)\s*\|\s*(\d+)\s*\|\s*([\d.]+%)\s*\|\s*(.*)', line)
                if match:
                    try:
                        id_atm = int(match.group(1).strip())
                        id_atm_str = f"{id_atm:08d}"  # Ensure ATM ID is 8 digits
                        nama_atm = match.group(2).strip()
                        jml_uang = match.group(3).strip()
                        percent = match.group(4).strip()
                        start_pagu = match.group(5).strip()
                        percent_value = float(percent.strip('%'))
                        if id_atm not in exceptions:
                            problem_details = (
                                f"saldo mendekati pagu dengan jumlah uang {jml_uang}, nilai tersebut {percent} dari total saldo, "
                                f"saldo mendekati pagu mulai pukul {start_pagu} pada ATM ID {id_atm_str}"
                            )
                            if percent_value > 10:
                                above_ten_percent.append({"ID_ATM": id_atm_str, "NAMA_ATM": nama_atm, "PROBLEM": problem_details, "START_TIME": start_pagu, "TYPE": error_type})
                            else:
                                problems.append({"ID_ATM": id_atm_str, "NAMA_ATM": nama_atm, "PROBLEM": problem_details, "START_TIME": start_pagu, "TYPE": error_type})
                    except ValueError:
                        print(f"Skipping line (ID_ATM not digit or malformed): {line.strip()}")
            elif atm_problem_section:
                match = re.match(r'\s*\d+\.\s*(\d+)\s*\|\s*([^|]+)\s*\|\s*(.*)\s*\|\s*(.*)', line)
                if match:
                    try:
                        id_atm = int(match.group(1).strip())
                        id_atm_str = f"{id_atm:08d}"  # Ensure ATM ID is 8 digits
                        nama_atm = match.group(2).strip()
                        start_error = match.group(3).strip()
                        ket = match.group(4).strip()
                        if id_atm not in exceptions:
                            problem_details = f"error dengan keterangan : ID ATM {id_atm_str} {ket} sejak jam {start_error}"
                            if "Reject Bin" in ket or "Currency Cassettes" in ket or "Receipt Paper" in ket:
                                error_type = 'Problem Supply Out'
                            problems.append({"ID_ATM": id_atm_str, "NAMA_ATM": nama_atm, "PROBLEM": problem_details, "START_TIME": start_error, "TYPE": error_type})
                    except ValueError:
                        print(f"Skipping line (ID_ATM not digit or malformed): {line.strip()}")
            else:
                match = re.match(r'\s*\d+\.\s*(\d+)\s*\|\s*([^|]+)\s*\|\s*(.*)', line)
                if match:
                    try:
                        id_atm = int(match.group(1).strip())
                        id_atm_str = f"{id_atm:08d}"  # Ensure ATM ID is 8 digits
                        nama_atm = match.group(2).strip()
                        if id_atm not in exceptions:
                            problem_details = match.group(3).strip()
                            problems.append({"ID_ATM": id_atm_str, "NAMA_ATM": nama_atm, "PROBLEM": problem_details, "TYPE": error_type})
                    except ValueError:
                        print(f"Skipping line (ID_ATM not digit or malformed): {line.strip()}")
    return problems, not_found, above_ten_percent

# Function to create messages and save to a new Excel file
def create_messages_and_save_to_excel(problems, not_found, above_ten_percent, atm_info_path, output_path):
    # Load the Excel data
    atm_info = pd.read_excel(atm_info_path, sheet_name='info')
    atm_info.columns = atm_info.columns.str.strip()  # Remove extra spaces in column names
    atm_info.columns = atm_info.columns.str.upper()  # Convert columns to uppercase to avoid case issues

    # Print the loaded ATM info data for debugging
    print("ATM Info Data:")
    print(atm_info.head())
    print("-" * 50)

    # Dictionaries to store the result data
    messages = defaultdict(list)
    new_history_records = []

    # Day names in Indonesian
    days_in_indonesian = {
        'Monday': 'Senin',
        'Tuesday': 'Selasa',
        'Wednesday': 'Rabu',
        'Thursday': 'Kamis',
        'Friday': 'Jumat',
        'Saturday': 'Sabtu',
        'Sunday': 'Minggu'
    }

    # Get current hour to determine greeting
    current_hour = datetime.now().hour
    if current_hour < 12:
        greeting = "Selamat Pagi"
    elif current_hour < 15:
        greeting = "Selamat Siang"
    else:
        greeting = "Selamat Sore"

    # Iterate through each problem and create the message text
    for problem in problems:
        if "ID_ATM" in problem:
            id_atm = problem["ID_ATM"]
            nama_atm = problem["NAMA_ATM"]
            problem_details = problem["PROBLEM"]
            start_time = problem.get("START_TIME", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
            error_type = problem["TYPE"]

            # Find the matching row in the Excel data
            match = atm_info[atm_info["ID_ATM"] == int(id_atm)]

            if not match.empty:
                nama_cabang = match.iloc[0]["NAMA_CABANG"]
                pic_name = match.iloc[0]["PIC_NAME"]
                phone = match.iloc[0]["PHONE"]

                # Append details to the message dictionary
                messages[nama_cabang].append({
                    "pic_name": pic_name,
                    "nama_atm": nama_atm,
                    "id_atm": id_atm,
                    "problem_details": problem_details,
                    "phone": phone,
                    "type": error_type
                })

                # Append new record to history
                now = datetime.now()
                day_name = days_in_indonesian[now.strftime("%A")]
                new_record = {
                    "TANGGAL INPUT": now.strftime('%d/%m/%Y %H:%M:%S'),
                    "HARI": day_name,
                    "TANGGAL": start_time.split(' ')[0],
                    "JAM": start_time.split(' ')[1],
                    "ID_ATM": id_atm,
                    "NAMA_ATM": nama_atm,
                    "PERMASALAHAN": problem_details,
                    "TINDAK LANJUT": "",
                    "KETERANGAN": "",
                    "3 HOURS": "",
                    "TYPE": error_type
                }
                new_history_records.append(new_record)
            else:
                print(f"No match found for ID_ATM {id_atm}")
                not_found.append({"ID_ATM": id_atm, "NAMA_ATM": nama_atm, "Problem Details": problem_details, "TYPE": error_type})
        elif "ATM_NAME" in problem:
            atm_name = problem["ATM_NAME"]
            problem_details = problem["PROBLEM"]
            start_time = problem.get("START_TIME", datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
            error_type = problem["TYPE"]

            # Find the matching row in the Excel data
            match = atm_info[atm_info["NAMA_ATM"] == atm_name]

            if not match.empty:
                nama_cabang = match.iloc[0]["NAMA_CABANG"]
                pic_name = match.iloc[0]["PIC_NAME"]
                phone = match.iloc[0]["PHONE"]

                # Append details to the message dictionary
                messages[nama_cabang].append({
                    "pic_name": pic_name,
                    "nama_atm": atm_name,
                    "id_atm": "",  # Assuming no ID available for ATM_NAME section
                    "problem_details": problem_details,
                    "phone": phone,
                    "type": error_type
                })

                # Append new record to history
                now = datetime.now()
                day_name = days_in_indonesian[now.strftime("%A")]
                new_record = {
                    "TANGGAL INPUT": now.strftime('%d/%m/%Y %H:%M:%S'),
                    "HARI": day_name,
                    "TANGGAL": start_time.split(' ')[0],
                    "JAM": start_time.split(' ')[1],
                    "ID_ATM": "",
                    "NAMA_ATM": atm_name,
                    "PERMASALAHAN": problem_details,
                    "TINDAK LANJUT": "",
                    "KETERANGAN": "",
                    "3 HOURS": "",
                    "TYPE": error_type
                }
                new_history_records.append(new_record)
            else:
                print(f"No match found for ATM_NAME {atm_name}")
                not_found.append({"ATM_NAME": atm_name, "Problem Details": problem_details, "TYPE": error_type})

    # Combine messages by cabang
    combined_messages = []
    for nama_cabang, details in messages.items():
        pic_name = details[0]['pic_name']
        phone = details[0]['phone']
        atm_details = ', '.join([f"{d['nama_atm']} ID {d['id_atm']}" for d in details])
        problem_details_combined = '\n\n'.join([d['problem_details'] for d in details])

        # Create the combined message text
        message = (
            f"{greeting},\n\n"
            f"Bapak/Ibu {pic_name},\n\n"
            f"Perkenalkan, saya Made Bramasta Vikana Putra, dari DJA pusat. Saya ingin memberitahukan bahwa ATM dengan details *{atm_details}* yang masih dalam kelolaan *{nama_cabang}* mendapatkan peringatan dengan rincian sebagai berikut:\n\n"
            f"{problem_details_combined}\n\n"
            "Mohon kesediaannya untuk segera menindaklanjuti permasalahan ini. \n"
            "Terima kasih atas perhatian dan kerjasamanya."
        )

        # URL encode the message
        encoded_message = urllib.parse.quote(message)

        # Create WhatsApp URL
        whatsapp_url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}"

        # Append the combined message to the list
        combined_messages.append({
            "NAMA_CABANG": nama_cabang,
            "Message": message,
            "PHONE": phone,
            "WhatsApp_URL": whatsapp_url,
            "TYPE": details[0]['type']
        })

    # Convert the results to a DataFrame
    results_df = pd.DataFrame(combined_messages)
    not_found_df = pd.DataFrame(not_found)
    above_ten_percent_df = pd.DataFrame(above_ten_percent)

    # Save the results to a new Excel file with multiple sheets
    with pd.ExcelWriter(output_path) as writer:
        results_df.to_excel(writer, sheet_name='Found', index=False)
        not_found_df.to_excel(writer, sheet_name='Not Found', index=False)
        above_ten_percent_df.to_excel(writer, sheet_name='Above 10 Percent', index=False)

    print(f"Messages saved to {output_path}")

    # Append new history records to history.xlsx
    history_df = pd.read_excel('history.xlsx')
    new_history_df = pd.DataFrame(new_history_records)
    updated_history_df = pd.concat([history_df, new_history_df], ignore_index=True)

    # Check if 3 hours have passed for each issue
    now = datetime.now()
    for i, row in updated_history_df.iterrows():
        if row['3 HOURS'] == "":
            issue_time = datetime.strptime(row['TANGGAL INPUT'], '%d/%m/%Y %H:%M:%S')
            if (now - issue_time).total_seconds() > 3 * 3600:
                updated_history_df.at[i, '3 HOURS'] = 'yes'

    updated_history_df.to_excel('history.xlsx', index=False)
    print("History updated successfully.")

# Main function to run the script
def main():
    text_file_path = 'report.txt'  # Path to the text file
    atm_info_path = 'info.xlsx'  # Path to the Excel file
    output_path = 'atm_problem_messages.xlsx'  # Path to the output Excel file

    # Load the Excel data
    atm_info = pd.read_excel(atm_info_path, sheet_name='info')
    exceptions = pd.read_excel(atm_info_path, sheet_name='exception')['ID_ATM'].tolist()
    
    atm_info.columns = atm_info.columns.str.strip()  # Remove extra spaces in column names
    atm_info.columns = atm_info.columns.str.upper()  # Convert columns to uppercase to avoid case issues
    
    problems, not_found, above_ten_percent = process_text_file(text_file_path, atm_info, exceptions)
    print(f"Parsed problems: {problems}")
    print(f"Not found: {not_found}")
    create_messages_and_save_to_excel(problems, not_found, above_ten_percent, atm_info_path, output_path)

if __name__ == "__main__":
    main()
