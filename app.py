import dash_bootstrap_components as dbc
import pandas as pd
from dash import html
from dash import dcc
from dash.dependencies import Input, Output, State
import plotly.express as px
import dash
from dash import dash_table
import Oneline
import Finalize_RMR_Data


first_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div(id='corp-eng-control-tabs', className='control-tabs', children=[
                dcc.Tabs(colors={'border': '#d6d6d6', 'primary': '#dc8633'},
                    style={'height':'50px', 'width': '100%', 'fontSize':12},
                    id='corp-eng-tabs', value='what-is',
                    children=[
                        dcc.Tab(
                            label='Home',
                            value='what-is',
                            children=html.Div(className='control-tab', children=[
                                html.Br(),
                                html.P(
                                    "The adjacent tabs contain tools to be used by the Corporate Engineering Team",
                                    style={"font-weight": "bold", "fontSize":15}
                                ),
                                html.Br(),
                                html.P(
                                    """
                                    Oneline: Excel formatted Aries Oneline
                                    """,
                                    style={"fontSize":14}
                                ),
                                html.P(
                                    """
                                    RMR: Reserve Management Record Input/Sign Off Sheets
                                    """,
                                    style={"fontSize":14}
                                ),
                                html.P(
                                    """
                                    Wiki: Information to access specific templates/files/queries/etc.
                                    """,
                                    style={"fontSize":14}
                                ),
                                html.Br(),
                                html.Br(),
                            ])
                        ),
                        dcc.Tab(
                            label='Oneline',
                            value='oneline-select',
                            children=html.Div(className='control-tab', children=[
                                html.Br(),
                                html.Label('Enter Scenario', style={"font-weight": "bold", "fontSize":14}),
                                dbc.Input(id='idscenario', type='text', value='EAP_Q321_SEC'),
                                html.Br(),
                                html.Label('Save as', style={"font-weight": "bold", "fontSize":14}),
                                dbc.Input(id='idsaveas', type='text'),
                                html.Br(),
                                html.Br(),
                                html.Button("Create Oneline Report", id='oneline-btn', n_clicks=0),
                                html.Br(),
                                html.Br(),
                            ])
                        ),
                        dcc.Tab(
                            label='RMR',
                            value='rmr-select',
                            children=html.Div(className='control-tab', children=[
                                html.Div(className='app-controls-block', children=[
                                    html.Div(
                                        className='fullwidth-app-controls-name',
                                        children="Select RMR Phase", style={"font-weight": "bold", "fontSize":14},
                                    ),
                                    dcc.Dropdown(
                                        id='rmr-dropdown',
                                        options=[
                                            {
                                                'label': 'RMR Email',
                                                'value': 'rmr-email'
                                            },
                                            {
                                                'label': 'RMR Check Status',
                                                'value': 'rmr-check-status'
                                            },
                                            {
                                                'label': 'RMR Finalize',
                                                'value': 'rmr-finalize'
                                            },
                                        ],
                                        value='dataset3',
                                    )
                                ]),
                            ])
                        ),
                        dcc.Tab(
                            label='Wiki',
                            value='wiki-select',
                            children=html.Div(className='control-tab', children=[
                                html.Div(className='app-controls-block', children=[
                                    html.Div(
                                        className='fullwidth-app-controls-name',
                                        children="Corporate Engineering Wiki", style={"font-weight": "bold", "fontSize":14},
                                    ),
                                ]),
                            ])
                        ),
                    ],
                ),
            ])
        ]
    )
)


second_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div(
                id='second-empty'
            ),

            html.Div(
                id='oneline-body',
                children=[
                    html.Label('Select Database', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='iddatabase',
                                 options=[
                                     {'label': 'ARIES_RSV', 'value': 'ARIES_RSV'},
                                     {'label': 'CORP_ENG_DATA_TOOLS', 'value': 'CORP_ENG_DATA_TOOLS'}
                                 ],
                                 value='ARIES_RSV'),
                    html.Br(),
                    html.Label('Enter Oneline Table', style={"font-weight": "bold"}),
                    dbc.Input(id='idaconeline', value='AC_ONELINE', type='text'),
                    html.Br(),
                    html.Label('Enter Property Table', style={"font-weight": "bold"}),
                    dbc.Input(id='idacproperty', value='AC_PROPERTY', type='text'),
                    html.Br(),
                    html.Label('Select Effective Date, Month', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='ideffmonth',
                                 options=[
                                     {'label': 'January', 'value': '01'},
                                     {'label': 'February', 'value': '02'},
                                     {'label': 'March', 'value': '03'},
                                     {'label': 'April', 'value': '04'},
                                     {'label': 'May', 'value': '05'},
                                     {'label': 'June', 'value': '06'},
                                     {'label': 'July', 'value': '07'},
                                     {'label': 'August', 'value': '08'},
                                     {'label': 'September', 'value': '09'},
                                     {'label': 'October', 'value': '10'},
                                     {'label': 'November', 'value': '11'},
                                     {'label': 'December', 'value': '12'}
                                 ],
                                 value='01'),
                    html.Br(),
                    html.Label('Select Effective Date, Year', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='ideffyear',
                                 options=[
                                     {'label': '2021', 'value': '2021'},
                                     {'label': '2022', 'value': '2022'},
                                     {'label': '2023', 'value': '2023'},
                                     {'label': '2024', 'value': '2024'}
                                 ],
                                 value='2021'),
                    html.Br(),
                    html.Label('User', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='iduser',
                                 options=[
                                     {'label': 'Anangela', 'value': 'agonzalez'},
                                     {'label': 'Assiya', 'value': 'abekniyazova'},
                                     {'label': 'Dominique', 'value': 'dfdearagao'},
                                     {'label': 'Joshua', 'value': 'jddearagao'},
                                     {'label': 'Leslie', 'value': 'larmentrout'}
                                 ],
                                 value='jddearagao'),
                    html.Br(),
                    html.Div(id='create_oneline', children=None),
                ]
            ),

            html.Div(
                id='rmr-email-body',
                children=[
                    html.Label('Quarter', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-quarter',
                                 options=[
                                     {'label': 'Q1', 'value': 'Q1'},
                                     {'label': 'Q2', 'value': 'Q2'},
                                     {'label': 'Q3', 'value': 'Q3'},
                                     {'label': 'Q4', 'value': 'Q4'},
                                 ]
                                 ),
                    html.Br(),
                    html.Label('Year', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-year',
                                 options=[
                                     {'label': '2021', 'value': '2021'},
                                     {'label': '2022', 'value': '2022'},
                                     {'label': '2023', 'value': '2023'},
                                     {'label': '2024', 'value': '2024'},
                                 ]
                                 ),
                    html.Br(),
                    html.Label('Enter Email Subject', style={"font-weight": "bold"}),
                    dbc.Input(id='id-email-subject', value='', type='text'),
                    html.Br(),
                    html.Label('Enter Email Body', style={"font-weight": "bold"}),
                    dbc.Textarea(id='id-email-body', style={'height': 200}),
                        ]
                    ),

            html.Div(
                id='rmr-final',
                children=[
                    html.Label('Quarter', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-quarter-final',
                                 options=[
                                     {'label': 'Q1', 'value': 'Q1'},
                                     {'label': 'Q2', 'value': 'Q2'},
                                     {'label': 'Q3', 'value': 'Q3'},
                                     {'label': 'Q4', 'value': 'Q4'},
                                 ],
                                 value='Q3',
                                 ),
                    html.Br(),
                    html.Label('Year', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-year-final',
                                 options=[
                                     {'label': '2021', 'value': '2021'},
                                     {'label': '2022', 'value': '2022'},
                                     {'label': '2023', 'value': '2023'},
                                     {'label': '2024', 'value': '2024'},
                                 ],
                                 value='2023',
                                 ),
                    html.Br(),
                    html.Label('Enter Email Subject', style={"font-weight": "bold"}),
                    dbc.Input(id='id-email-subject-final', value='', type='text'),
                    html.Br(),
                    html.Label('Enter Email Body', style={"font-weight": "bold"}),
                    dbc.Textarea(id='id-email-body-final', style={'height': 200}),
                    html.Br(),
                    dbc.Button('Finalize RMR', id='finalize-btn', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    dbc.Button('Create Aries csv Files and Upload to Aries', id='update-aries-btn', color='success', n_clicks=0, outline=True),
                ]
                    )
        ]
    )
)


third_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div(
                id='rmr-email-dd',
                children=[
                    html.Label('Drilling & Completions', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-dc',
                                 options=[
                                     {'label': 'Lance', 'value': 'Q1'},
                                     {'label': 'Daniel', 'value': 'Q2'},
                                     {'label': 'Corporate Engineering', 'value': 'Q3'}
                                 ],
                                 value=''),
                    html.Br(),
                    dbc.Button('Send Email to D&C', id='email-dc', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('GIS', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-gis',
                                 options=[
                                     {'label': 'Maria', 'value': 'Q1'},
                                     {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to GIS', id='email-gis', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Land Administration', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-land-admin',
                                     options=[
                                         {'label': 'Betty', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Land Administration', id='email-land-admin', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Land', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-land',
                                     options=[
                                         {'label': 'Matt', 'value': 'Q1'},
                                         {'label': 'Tanner', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Land', id='email-land', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Gas Marketing', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-gas-mark',
                                     options=[
                                         {'label': 'Megan', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Gas Marketing', id='email-gas-mark', color='success', n_clicks=0, outline=True),
                ]
            ),

            html.Div(
                # id='rmr-final-capex',
                # children=[
                    # html.Div(id='wet-drilling-capex')
                    dcc.Graph(id='wet-drilling-capex')
                # ]
            )
        ]
    )
)


fourth_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div(
                id='rmr-email-btn',
                children=[
                    html.Label('Liquids Marketing', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='id-liq-mark',
                                 options=[
                                     {'label': 'Kelsey', 'value': 'Q1'},
                                     {'label': 'Corporate Engineering', 'value': 'Q3'}
                                 ],
                                 value=''),
                    html.Br(),
                    dbc.Button('Send Email to Liquids Marketing', id='email-liq-mark', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Planning', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-rep',
                                     options=[
                                         {'label': 'Sara', 'value': 'Q1'},
                                         {'label': 'Mary', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Planning', id='emailplanning', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Sub Surface', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-subsurf',
                                     options=[
                                         {'label': 'Chris', 'value': 'Q1'},
                                         {'label': 'Connor', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Sub Surface', id='email-sub-surf', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Midstream', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-midstream',
                                     options=[
                                         {'label': 'Eric', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Midstream', id='email-mid', color='success', n_clicks=0, outline=True),
                    html.Br(),
                    html.Br(),
                    html.Label('Production Operations', style={"font-weight": "bold"}),
                        dcc.Dropdown(id='id-prod-op',
                                     options=[
                                         {'label': 'Brian', 'value': 'Q1'},
                                         {'label': 'Jackie', 'value': 'Q1'},
                                         {'label': 'Corporate Engineering', 'value': 'Q3'}
                                     ],
                                     value=''),
                    html.Br(),
                    dbc.Button('Send Email to Prod Operations', id='email-prod', color='success', n_clicks=0, outline=True),
                ]
            )
        ]
    )
)

navbar = dbc.Navbar(
    dbc.Container(
        [
            html.A(
                dbc.Row(
                    [
                        dbc.Col(html.Img(src='/assets/Encino-logo.png', height="30px")),
                        dbc.Col(dbc.NavbarBrand("Encino Corporate Engineering", className="ms-2")),
                    ],
                    align="left",
                    className="g-0",
                ),
                href="https://www.encinoenergy.com/",
                style={"textDecoration": "none"},
            ),
            dbc.NavbarToggler(id="navbar-toggler", n_clicks=0),
            dbc.Row(
                [
                    dbc.Col(html.Img(src='/assets/Sharepoint.png', height="30px"))
                ],
                align='right',
            )
        ]
    ),
    color="#3e4827",
    dark=True,
)

cards = dbc.Row(
    [
        dbc.Col(first_card, width=3),
        dbc.Col(second_card, width=4),
        dbc.Col(third_card, width=2),
        dbc.Col(fourth_card, width=2),
    ]
)

app = dash.Dash(external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = 'Encino Corporate Engineering App'
app.layout = html.Div(
                    [navbar, cards],
                    style={'background': '#f6f6f6'})


@app.callback(
    Output('oneline-body', 'style'),
    [Input('corp-eng-tabs', 'value')]
)
def show_hide_custom_colorscheme(oneline_data):
    return {'display': 'block' if oneline_data == 'oneline-select' else 'none'}


@app.callback(
    Output("create_oneline", "children"),
    [Input("oneline-btn", "n_clicks")],
    [State('iddatabase', 'value'), State('idaconeline', 'value'), State('idacproperty', 'value'), State('idscenario', 'value'),
     State('ideffmonth', 'value'), State('ideffyear', 'value'), State('iduser', 'value'), State('idsaveas', 'value')]
)
def buttons(onelinebtn, database, aconeline, acproperty, scenario, month, year, user, saveas):
    if onelinebtn:
        filepath = Oneline.oneline_report(database, aconeline, acproperty, scenario, month, year, user, saveas)
        return filepath


@app.callback(
    Output('rmr-email-body', 'style'),
    [Input('rmr-dropdown', 'value')]
)
def show_hide_custom_colorscheme(rmr_email):
    return {'display': 'block' if rmr_email == 'rmr-email' else 'none'}


@app.callback(
    Output('rmr-email-dd', 'style'),
    [Input('rmr-dropdown', 'value')]
)
def show_hide_custom_colorscheme(rmr_email):
    return {'display': 'block' if rmr_email == 'rmr-email' else 'none'}


@app.callback(
    Output('rmr-email-btn', 'style'),
    [Input('rmr-dropdown', 'value')]
)
def show_hide_custom_colorscheme(rmr_email):
    return {'display': 'block' if rmr_email == 'rmr-email' else 'none'}


@app.callback(
    Output('rmr-final', 'style'),
    [Input('rmr-dropdown', 'value')]
)
def show_hide_custom_colorscheme(rmr_email):
    return {'display': 'block' if rmr_email == 'rmr-finalize' else 'none'}


@app.callback(
    Output('rmr-final-capex', 'style'),
    [Input('rmr-dropdown', 'value')]
)
def show_hide_custom_colorscheme(rmr_email):
    return {'display': 'block' if rmr_email == 'rmr-finalize' else 'none'}


@app.callback(
    Output("wet-drilling-capex", "children"),
    [Input("update-aries-btn", "n_clicks")],
    [State('id-year-final', 'value'), State('id-quarter-final', 'value')]
)
def buttons(wet_drill_capex, year, quarter):
    if wet_drill_capex:
        df_d_wet, df_d_dry, df_c_wet, df_c_dry = Finalize_RMR_Data.aries_csv(year, quarter)
        df_d_wet = df_d_wet[['Lateral_Length', 'drill_wet']]
        fig_wet_drill = px.scatter(df_d_wet, x="Lateral_Length", y="drill_wet")
        data = df_d_wet.to_dict('rows')
        columns = [{"name": i, "id": i} for i in (df_d_wet.columns)]
        print(fig_wet_drill)
        print(type(fig_wet_drill))
        return fig_wet_drill
        # return dash_table.DataTable(data=data, columns=columns)


if __name__ == '__main__':
    app.run_server(debug=True)
