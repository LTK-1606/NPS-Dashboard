import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import webview
import threading
import time
import os

# Initialize the Dash app
app = dash.Dash(__name__)

# Function to find the specifically named Excel file in the current directory
def find_specific_excel_file(filename):
    for file in os.listdir('.'):
        if file == filename:
            return file
    return None

# Specify the name of the Excel file to search for
excel_filename = 'NPS Quarterly Summary.xlsx'

# Find the specified Excel file in the current directory
excel_file_path = find_specific_excel_file(excel_filename)

# Check if the file was found
if excel_file_path is None:
    raise ValueError(f"Excel file '{excel_filename}' not found in the current directory.")

# Load the Excel file
excel_data = pd.ExcelFile(excel_file_path)
sheet_names = excel_data.sheet_names

# App layout
app.layout = html.Div([
    dcc.Dropdown(
        id='input-sheet-name',
        options=[{'label': sheet, 'value': sheet} for sheet in sheet_names],
        value=sheet_names[0],
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
    app.run_server(debug=True, port=8050, use_reloader=False)

# Function to create webview window after a delay
def create_webview_with_delay(delay_seconds):
    time.sleep(delay_seconds)  # Delay for specified seconds
    webview.create_window("Dash App", "http://127.0.0.1:8050/", width=1200, height=1600, resizable=True)
    webview.start()

# Entry point of the script
if __name__ == '__main__':
    dash_thread = threading.Thread(target=run_dash)
    dash_thread.start()

    # Delay webview creation to ensure server is up
    create_webview_with_delay(delay_seconds=0)  # Adjust delay as needed
