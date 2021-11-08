import dash
from dash import dcc
import dash_bootstrap_components as dbc
from dash import html
from dash.dependencies import Input, Output
import plotly.express as px

import pandas as pd

df = pd.read_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/CAPEX.csv')

app = dash.Dash(__name__)

first_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div('Testing'),

            html.Div(
                dbc.Button('Test Button', id='test-btn', color='success', n_clicks=0, outline=True)
            ),

            html.Div([
                dcc.Graph(id='graph-test')
            ]
            )
        ]
    )
)


cards = dbc.Row(
    [
        dbc.Col(first_card, width=3)]
)

app.layout = html.Div([cards])


@app.callback(
    Output('graph-test', 'figure'),
    Input('test-btn', 'n-clicks'))
def update_figure(test_btn):
    print(test_btn)
    if test_btn:
        fig = px.scatter(df, x="Lateral_Length", y="compl_wet")
        print(fig)
        print(type(fig))
        return fig


if __name__ == '__main__':
    app.run_server(debug=True)
