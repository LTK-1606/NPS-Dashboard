import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import webview
import threading
import time
import os
import sys
from datetime import datetime
from screeninfo import get_monitors

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

# Initialize the Dash app and Create Excel File
app = dash.Dash(__name__)
main()

# Function to find the specifically named Excel file in the script directory
def find_specific_excel_file(script_dir, filename):
    excel_file_path = os.path.join(script_dir, filename)
    if os.path.isfile(excel_file_path):
        return excel_file_path
    else:
        return None

# Determine script directory
if getattr(sys, 'frozen', False):
    # When running as a bundled executable (e.g., PyInstaller)
    script_dir = os.path.dirname(sys.executable)
else:
    # When running as a script
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Specify the name of the Excel file to search for
excel_filename = 'NPS Quarterly Summary.xlsx'

# Find the specified Excel file in the script directory
excel_file_path = find_specific_excel_file(script_dir, excel_filename)

# Check if the file was found
if excel_file_path is None:
    raise ValueError(f"Excel file '{excel_filename}' not found in the script directory: {script_dir}")

# Load the Excel file
excel_data = pd.ExcelFile(excel_file_path)
sheet_names = excel_data.sheet_names
displayed_sheet_names = sheet_names[:-3]

# App layout
app.layout = html.Div([
    dcc.Dropdown(
        id='input-sheet-name',
        options=[{'label': sheet, 'value': sheet} for sheet in displayed_sheet_names],
        value=displayed_sheet_names[0],
        clearable=False,  # Prevent clearing the dropdown
        placeholder="Select a sheet"
    ),
    html.Div(id='output-data'),
    html.Div(id='output-weighted-scores')
])

# Function to calculate weighted scores
def calculate_weighted_scores(df):
    # Example weights (adjust according to your data)
    weights = {'1': 1, '2': 2, '3': 3, '4': 4, '5': 5}
    
    # Initialize an empty list to store weighted scores
    weighted_scores = []

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        # Calculate weighted score for the current row based on available numeric columns
        weighted_score = 0
        for col_name in df.columns[1:]:
            if pd.api.types.is_numeric_dtype(df[col_name]):  # Check if column is numeric
                weighted_score += row[col_name] * weights.get(str(col_name), 0)

        # Append the weighted score to the list
        weighted_scores.append(weighted_score)
    
    # Add the weighted scores list as a new column to the DataFrame
    df['Weighted_Score'] = weighted_scores
    
    return df

# Callback to update the data based on the input sheet name
@app.callback(
    [Output('output-data', 'children'),
     Output('output-weighted-scores', 'children')],
    [Input('input-sheet-name', 'value')]
)
def update_output(sheet_name):
    try:
        # Load the specified sheet into a DataFrame
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        # Check if the DataFrame is empty
        if df.empty:
            return html.Div([
                html.H4(f"No Data to Display for {sheet_name}", style={'textAlign': 'center', 'fontSize': '24px'})
            ]), html.Div()
        
        # Select only numeric columns and calculate total
        numeric_df = df.select_dtypes(include=['number'])
        df['Total'] = numeric_df.sum(axis=1)

        # Annotations for Total Reviews
        annotations1 = [
            {
                'x': x,
                'y': y + 0.05 * max(df['Total']),  # Adjust y to place annotations above bars
                'text': str(y),
                'xref': 'x',
                'yref': 'y',
                'showarrow': False,
                'font': {'size': 10},
                'align': 'center'
            }
            for x, y in zip(df[df.columns[0]], df['Total'])
        ]

        # Calculate weighted scores
        df_weighted_scores = calculate_weighted_scores(df)

        # Annotations for Weighted Scores
        annotations2 = [
            {
                'x': x,
                'y': y + 0.05 * max(df_weighted_scores['Weighted_Score']),  # Adjust y for annotations
                'text': str(y),
                'xref': 'x',
                'yref': 'y',
                'showarrow': False,
                'font': {'size': 10},
                'align': 'center'
            }
            for x, y in zip(df[df.columns[0]], df_weighted_scores['Weighted_Score'])
        ]

        # Create figure for total reviews bar chart
        fig_total_reviews = {
            'data': [
                {'x': df[df.columns[0]], 'y': df['Total'], 'type': 'bar', 'name': 'Total Reviews'},
            ],
            'layout': {
                'title': {
                    'text': f"Total Reviews for {sheet_name}",
                    'font': {'size': 24}
                },
                'xaxis': {
                    'title': 'Sales Executives',  # X-axis title
                    'tickmode': 'linear',
                    'tick0': 0,
                    'dtick': 1  # X-axis tick interval
                },
                'yaxis': {
                    'title': 'Total Reviews',  # Y-axis title
                    'tickmode': 'linear',
                    'tick0': 0,
                    'dtick': 10  # Y-axis tick interval
                },
                'annotations': annotations1  # Annotations for total reviews
            }
        }

        # Create figure for weighted scores bar chart
        fig_weighted_scores = {
            'data': [
                {'x': df[df.columns[0]], 'y': df_weighted_scores['Weighted_Score'], 'type': 'bar', 'name': 'Weighted Scores'},
            ],
            'layout': {
                'title': {
                    'text': f"Weighted Scores for {sheet_name}",
                    'font': {'size': 24}
                },
                'xaxis': {
                    'title': 'Sales Executives',  # X-axis title
                    'tickmode': 'linear',
                    'tick0': 0,
                    'dtick': 1  # X-axis tick interval
                },
                'yaxis': {
                    'title': 'Weighted Scores',  # Y-axis title
                    'tickmode': 'linear',
                    'tick0': 0,
                    'dtick': 50  # Y-axis tick interval
                },
                'annotations': annotations2  # Annotations for weighted scores
            }
        }

        return (
            html.Div([
                html.H4(f"Data from {sheet_name}", style={'textAlign': 'center', 'fontSize': '24px'}),
                dcc.Graph(id='total-reviews', figure=fig_total_reviews)
            ]),
            dcc.Graph(id='weighted-scores', figure=fig_weighted_scores)
        )
    except Exception as e:
        return html.Div([
            html.H4(f"Error: {str(e)}", style={'textAlign': 'center', 'fontSize': '24px'})
        ]), html.Div()

# Function to start Dash server
def run_dash():
    app.run_server(debug=False, port=8050, use_reloader=False)

# Function to create webview window after a delay
def create_webview_with_delay(delay_seconds):
    time.sleep(delay_seconds)  # Delay for specified seconds
    width=screen_width,
    height=screen_height,
    webview.create_window("Dash App", "http://127.0.0.1:8050/", width=screen_width, height=screen_height, resizable=True)
    webview.start()

monitor = get_monitors()[0]  # Assumes single monitor setup
screen_width = monitor.width
screen_height = monitor.height

# Entry point of the script
if __name__ == '__main__':
    dash_thread = threading.Thread(target=run_dash)
    dash_thread.start()

    # Delay webview creation to ensure server is up
    create_webview_with_delay(delay_seconds=0)  # Adjust delay as needed
