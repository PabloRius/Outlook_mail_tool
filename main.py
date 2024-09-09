import dash
import base64
import io
import pandas as pd
import plotly.express as px
from dash import dcc, html, Input, Output, State
from dash.exceptions import PreventUpdate

from outlook_parser import CSVMailReader, PSTMailReader

app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Mail Reader Web Interface - PGR"),

    html.Div([
        dcc.Upload(
            id='upload-data',
            children=html.Div([
                'Drag and Drop or ',
                html.A('Select a PST File')
            ]),
            style={
                'width': '100%',
                'height': '60px',
                'lineHeight': '60px',
                'borderWidth': '1px',
                'borderStyle': 'dashed',
                'borderRadius': '5px',
                'textAlign': 'center',
                'margin': '10px',
                'textDecoration': 'underline',
                'cursor': 'pointer'
            },
            multiple=False
        ),
        html.Div(id='file-upload-status')
    ]),

    dcc.Store(id='stored-data'),

    html.Div([
        dcc.Dropdown(
            id="time-period",
            options=[
                {'label': 'Per Month', 'value':'M'},
                {'label': 'Per Week', 'value':'W'},
                {'label': 'Per Day', 'value':'D'},
            ],
            value='M',
            clearable=False
        ),

        dcc.Graph(id='time-series-graph'),

        dcc.Graph(id='average-emails-graph'),

        dcc.Graph(id='after-hours-emails-graph')
    ])
])

def parse_contents(contents, filename):
    """Parse the file uploaded by the user"""
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename.lower():
            df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
        else:
            return html.Div(['Unsupported file format. Please upload a CSV file.'])
    except Exception as e:
        return html.Div([f'There was an error processing this file: {e}'])

    return df

@app.callback(
    Output('file-upload-status', 'children'),
    Output('stored-data', 'data'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename'),
)
def update_output(file_contents, filename):
    if file_contents is None:
        raise PreventUpdate

    try:
        content_type, content_string = file_contents.split(",")
        decoded = base64.b64decode(content_string)
        file = io.BytesIO(decoded)

        mail_reader = PSTMailReader(file)

        df = mail_reader.df
        
        df_json = mail_reader.df.to_json(date_format='iso', orient='split')

        return f'File "{filename}" uploaded successfully!', df_json

    except Exception as e:
        return f'File {filename} failed to upload: {e}', None

@app.callback(
    Output('time-series-graph', 'figure'),
    Output('average-emails-graph', 'figure'),
    Output('after-hours-emails-graph', 'figure'),
    Input('stored-data', 'data'),
    Input('time-period', 'value')
)
def update_graphs(data, time_period):
    if data is None:
        raise PreventUpdate

    df = pd.read_json(data, orient='split')

    df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
    df = df.dropna(subset=["Date"])

    df_grouped = df.set_index("Date").resample(time_period).size().reset_index(name="Count")
    df_grouped["Date"] = pd.to_datetime(df_grouped["Date"])

    max_count = df_grouped["Count"].max()
    yaxis_upper_limit = max(max_count + 10, max(20, min(100, max_count * 1.2)))

    period_label = {"M": "Month", "W": "Week", "D": "Day"}[time_period]

    fig_emails = px.line(df_grouped, x='Date', y='Count', title=f'Emails Received Per {period_label}')
    fig_emails.update_yaxes(range=[0, yaxis_upper_limit]) 

    if time_period == 'M':
        avg_count = df_grouped["Count"].mean()

        df_avg = pd.DataFrame({
            'Interval': [period_label],
            'Average Count': [avg_count]
        })

        fig_avg = px.bar(df_avg, x='Interval', y='Average Count', title=f'Average Emails Received Per {period_label}')
        fig_avg.update_yaxes(range=[0, yaxis_upper_limit])

    else:
        df_grouped['Month'] = df_grouped['Date'].dt.to_period('M').astype(str)

        df_avg = df_grouped.groupby('Month')['Count'].mean().reset_index()

        fig_avg = px.bar(df_avg, x='Month', y='Count', title=f'Average Emails Received Per {period_label} in each Month')
        fig_avg.update_yaxes(range=[0, yaxis_upper_limit])

    after_hours_time = pd.to_datetime("13:30").time()
    print(df["Date"])
    print(df["Date"].dt.time)

    df_after_hours = df[df["Date"].dt.time > after_hours_time]
    df_after_hours_grouped = df_after_hours.resample('D', on="Date").size().reset_index(name="Count")

    fig_after_hours = px.bar(df_after_hours_grouped, x='Date', y='Count', title='Emails Received After 13:30 Per Day')
    fig_after_hours.update_yaxes(range=[0, df_after_hours_grouped["Count"].max() + 5])

    return fig_emails, fig_avg, fig_after_hours


if __name__ == '__main__':
    app.run_server(debug=True)
