import os
import pandas as pd
import sys
from datetime import datetime

#Reads Files from Same Directory
def read_files(script_dir):

    #Construct paths to Excel files
    df_sg_path = os.path.join(script_dir, 'C2B Customer Satisfaction Survey (Responses).xlsx')
    df_my_path = os.path.join(script_dir, 'MY C2B Customer Satisfaction Survey (Responses).xlsx')
    df_th_path = os.path.join(script_dir, 'TH C2B Customer Satisfaction Survey (Responses).xlsx')

    #Read Excel files into DataFrames
    try:
        print(f"Reading Excel files...")
        df_sg = pd.read_excel(df_sg_path)
        df_my = pd.read_excel(df_my_path)
        df_th = pd.read_excel(df_th_path)
        print(f"Excel files read successfully.")
    except FileNotFoundError as e:
        print(f"Error: Excel file not found at path: {e.filename}")
        sys.exit(1)

    df = [df_sg, df_my, df_th]
    return df

#Convert Months to Quarters
def month_to_quarter(month):
    if month in ['January', 'February', 'March']:
        return 1
    elif month in ['April', 'May', 'June']:
        return 2
    elif month in ['July', 'August', 'September']:
        return 3
    elif month in ['October', 'November', 'December']:
        return 4

#Data Processing
def process(data):
    #Filter to Desired Year
    year = datetime.now().year
    data = data[data['Timestamp'].dt.year == year]

    #Group by Month
    data['Month'] = pd.to_datetime(data['Timestamp'], format='%d/%m/%y', errors='coerce').dt.strftime('%B')
    
    #Group by Quarter
    data['Quarter'] = data['Month'].apply(month_to_quarter)

    return data

#Data Cleaning/Processing
def cleanup(df, column, separators = ['and', '&']):
    new_rows = []

    #Look for any Rows with Multiple Names
    for index, row in df.iterrows():
        names = [row[column]]
        for separator in separators:
            temp_names = []
            for name in names:
                temp_names.extend(name.split(f' {separator} '))
            names = [name.strip() for name in temp_names]

        #Split row into multiple entries for every name present
        if len(names) > 1:
            for name in names:
                new_row = row.copy()
                new_row[column] = name
                new_rows.append(new_row)
        else:
            new_rows.append(row)

    #Update the original DataFrame with the new rows
    df = pd.DataFrame(new_rows)
    return df

#Updating Name Column to be Uniform
def update_name(df, name_list, column):
    for index, row in df.iterrows():
        changed = False
        name = row[column]
        #Loop through Names to Find a Match
        for match in name_list:
            if match.lower() in name.lower():
                df.loc[index, column] = match
                changed = True
                break        
        #Update Entries with No Matches
        if changed == False:
            match = 'No Name'
            df.loc[index, column] = match

#Data Filtering
def filter(data_info):
    df = []
    for data, column, sales_list in data_info:
        #Data Processing/Cleaning
        pd.options.mode.chained_assignment = None  #Disable Warnings
        data = process(data)
        data[column] = data[column].astype(str)
        pd.options.mode.chained_assignment = 'warn' #Enable Warnings
        data = cleanup(data, column)   
        update_name(data, sales_list, column) 
        df.append(data)
    return df

#Creates Pivot Tables by Quarters
def create_pivot_table(df, value, index, column):
    quarters = [1, 2, 3, 4]
    pivot_tables = []

    for quarter in quarters:
        df_q = df[df['Quarter'] == quarter]
        pivot_table = pd.pivot_table(df_q, 
                                     values = value,
                                     index = index,
                                     columns = column,
                                     aggfunc='count',
                                     fill_value=0 )
        pivot_table = pivot_table.reset_index()
        pivot_tables.append(pivot_table)
    return pivot_tables

#Exports Pivot Tables into Individual Sheets
def export_pivot(pivot_list, country, writer):
    i = 0    
    while i < 4:
        sheet_name = f"{country} Q{i + 1}"
        pivot_list[i].to_excel(writer, sheet_name = sheet_name, index = False)
        i += 1

#Create New Excel File with Multiple Sheets
def create_excel(script_dir, sg_pivot, my_pivot, th_pivot, df_sg, df_my, df_th):
    output_file = 'NPS Quarterly Summary.xlsx'
    print(f"Writing data to Excel file: {output_file}...")
    output_path = os.path.join(script_dir, output_file)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer: 
        export_pivot(sg_pivot, "SG", writer)
        export_pivot(my_pivot, "MY", writer)
        export_pivot(th_pivot, "TH", writer)
        df_sg.to_excel(writer, sheet_name = "SG Raw", index = False)
        df_my.to_excel(writer, sheet_name = "MY Raw", index = False)
        df_th.to_excel(writer, sheet_name = "TH Raw", index = False)

    print(f"Excel file '{output_file}' successfully created in {script_dir}")

def main():
    
    #Determine script directory
    if getattr(sys, 'frozen', False):
        #When running as a bundled executable (e.g., PyInstaller)
        script_dir = os.path.dirname(sys.executable)
    else:
        #When running as a script
        script_dir = os.path.dirname(os.path.abspath(__file__))

    #Reads Files into Dataframes List
    df = read_files(script_dir)
    
    #List of Sales Team
    sg_sales = ['Jasmine', 'Zhengjun', 'Jun', 'Jezelle', 'Joanna', 'Berlyn', 
                'Elaine', 'Leng Kiat', 'Roger', 'Katherine', 'Sharon', 'Darryl',
                'Norfazlin', 'Emir', 'Peggy', 'Diana', "A'rif Alimi", 'Mann', 'Mashrurah', 
                'Adeyrah', 'Mel', 'Mark', 'Nurul Nadia', 'Lishan']
    my_sales = ['Sook Ling', 'Hisham', 'Jia', 'Nadzirah', 'Alaina', 'Adeline', 'Mel']       
    th_sales = ['Pitchapak (Guitar)', 'Nareenart (Toei)', 'Monsicha (Yok)', 
                'Duangcheewan (Kratai)', 'Punchita (Belle)',
                'Pattaratiyaporn (Gap)', 'Pasu (Au)', 
                'Nisarat (Earn)', 'Sittichok (Job)', 'Konkanok (Teen)']   
    
    #Data Info
    data_info = [
        (df[0], 'Your Motorist Client Sales Executive (this section has been pre-filled for you)', sg_sales),
        (df[1], 'Your Motorist Sales Executive (this section has been pre-filled for you)\nEksekutif Jualan Pemandu anda (bahagian ini telah dipraisi untuk anda)', my_sales),
        (df[2], 'เจ้าหน้าที่มอเตอริสต์ผู้ให้บริการ(ข้อมูลส่วนนี้ระบบกรอกอัตโนมัติให้คุณ)', th_sales)
    ]

    #Data Filtering
    df = filter(data_info)
    
    #SG Pivot Tables
    sg_pivot = create_pivot_table(df[0],
                                'Your Enquiry ID (this section has been pre-filled for you)',
                                'Your Motorist Client Sales Executive (this section has been pre-filled for you)',
                                'How was your experience with your Motorist Customer Representative?' )
    #MY Pivot Tables
    my_pivot = create_pivot_table(df[1],
                                'Your Enquiry ID (this section has been pre-filled for you)\nID Pertanyaan Anda (bahagian ini telah di pra-isi untuk anda)',
                                'Your Motorist Sales Executive (this section has been pre-filled for you)\nEksekutif Jualan Pemandu anda (bahagian ini telah dipraisi untuk anda)',
                                'How was your experience with your Motorist Customer Representative?\nBagaimanakah pengalaman anda dengan pegawai khidmat pelanggan Motorist?' )
    #TH Pivot Tables
    th_pivot = create_pivot_table(df[2],
                                'หมายเลขผู้ใช้บริการของคุณ (ข้อมูลส่วนนี้ระบบกรอกอัตโนมัติให้คุณ)',
                                'เจ้าหน้าที่มอเตอริสต์ผู้ให้บริการ(ข้อมูลส่วนนี้ระบบกรอกอัตโนมัติให้คุณ)',
                                'ระดับความพึงพอใจของท่านในการบริการของเจ้าหน้าที่มอเตอริสต์' )
    
    #Output New Excel File
    create_excel(script_dir, sg_pivot, my_pivot, th_pivot, df[0], df[1], df[2])

if __name__ == "__main__":
    main()