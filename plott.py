import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
import os

# Define the file name
file_name = 'Statistics.xlsx'

# Check if the file exists in the same directory as the script
if not os.path.exists(file_name):
    raise FileNotFoundError(f"File not found: {file_name}")

# Define a mapping of languages to their corresponding sheet names
sheet_languages = {
    'English': 'English',
    'Hindi': 'Hindi',
    'Kannada': 'Kannada',
    'Tamil': 'Tamil',
    'Telugu': 'Telugu',
    'Hinglish': 'Hinglish',
    'Malayalam': 'Malayalam',
    'Gujarati': 'Gujarati',
    'Marathi': 'Marathi',
    'Bengali': 'Bengali'
}

# Initialize Dash app
app = Dash(__name__)

# Define layout
app.layout = html.Div([
    html.Div([
        html.H1("Remark Statistics", style={'textAlign': 'center'}),
        html.Div([
            html.Div([
                html.Label("Language:", style={'margin-right': '10px'}),
                dcc.Dropdown(
                    id='language-dropdown',
                    options=[{'label': lang, 'value': lang} for lang in sheet_languages.keys()],
                    value='English',
                    style={'width': '150px'}
                ),
            ], style={'margin-bottom': '20px', 'display': 'inline-block', 'verticalAlign': 'top'}),
            html.Div([
                html.Label("Model:", style={'margin-right': '10px'}),
                dcc.Dropdown(id='model-dropdown', style={'width': '200px'}),
            ], style={'margin-bottom': '20px', 'display': 'inline-block', 'verticalAlign': 'top'}),
            html.Div([
                html.Label("Remark:", style={'margin-right': '10px'}),
                dcc.Dropdown(id='remark-dropdown', style={'width': '150px'}),
            ], style={'margin-bottom': '20px', 'display': 'inline-block', 'verticalAlign': 'top'}),
        ], style={'textAlign': 'center', 'backgroundColor': '#C8E6C9', 'padding': '20px', 'borderRadius': '10px'}),
    ], style={'backgroundColor': '#A5D6A7', 'padding': '20px', 'borderRadius': '10px', 'marginBottom': '20px'}),
    dcc.Graph(id='remark-graph', style={'height': '700px', 'width': '80%', 'margin': 'auto'})
], style={'backgroundColor': '#E8F5E9', 'padding': '20px'})

# Callback to update model dropdown based on language selection
@app.callback(
    Output('model-dropdown', 'options'),
    Input('language-dropdown', 'value')
)
def set_models(selected_language):
    try:
        df = pd.read_excel(file_name, sheet_name=sheet_languages[selected_language])
        models = {col.split(' - ')[0] for col in df.columns if ' - ' in col}
        return [{'label': model, 'value': model} for model in models]
    except Exception as e:
        print(f"Error in set_models: {e}")
        return []

# Callback to update remark dropdown based on language and model selection
@app.callback(
    Output('remark-dropdown', 'options'),
    Input('language-dropdown', 'value'),
    Input('model-dropdown', 'value')
)
def set_remarks(selected_language, selected_model):
    try:
        df = pd.read_excel(file_name, sheet_name=sheet_languages[selected_language])
        if selected_model:
            remark_columns = [col for col in df.columns if col.startswith(f'{selected_model} - Remark')]
            return [{'label': col.split(' - ')[-1], 'value': col} for col in remark_columns]
        return []
    except Exception as e:
        print(f"Error in set_remarks: {e}")
        return []

# Callback to update graph based on language, model, and remark selection
@app.callback(
    Output('remark-graph', 'figure'),
    Input('language-dropdown', 'value'),
    Input('model-dropdown', 'value'),
    Input('remark-dropdown', 'value')
)
def update_graph(selected_language, selected_model, selected_remark):
    try:
        if selected_language and selected_model and selected_remark:
            df = pd.read_excel(file_name, sheet_name=sheet_languages[selected_language])
            selected_column = selected_remark
            percentage_column = selected_column.replace('Remark', 'Percentage')

            if selected_column in df.columns and percentage_column in df.columns:
                data = df[[selected_column, percentage_column]].dropna()
                data.columns = ['Remarks', 'Percentage']

                grouped_df = data.groupby('Remarks')['Percentage'].mean().reset_index()

                fig = px.bar(
                    grouped_df, x='Remarks', y='Percentage', 
                    title=f'Percentage of Remarks - {selected_language} - {selected_model}',
                    color_discrete_sequence=['#1f77b4']  # Blue color
                )
                fig.update_traces(
                    hovertemplate='%{y:.2f}%<br>Remark: %{x}',
                    textposition='none'  # Remove the text from being displayed directly on the graph
                )
                fig.update_layout(
                    xaxis_title='Remarks', yaxis_title='Percentage',
                    margin=dict(l=40, r=40, t=80, b=40),
                    plot_bgcolor='#E8F5E9',
                    hovermode='x unified',
                    xaxis=dict(showticklabels=False)  # Hide the x-axis tick labels
                )
                return fig
            else:
                return {}
        else:
            return {}
    except Exception as e:
        print(f"Error in update_graph: {e}")
        return {}

# Run the app
if __name__ == '__main__':
    app.run_server(port=5000, debug=True)
