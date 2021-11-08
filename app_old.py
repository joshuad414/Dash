import base64

import dash_bootstrap_components as dbc
from dash import html
from dash import dcc
from dash.dependencies import Input, Output, State
import Oneline

from layout_helper import run_standalone_app


def header_colors():
    return {
        'bg_color': '#3e4827',
        'font_color': 'white',
    }


test_card = dbc.Card(
    dbc.CardBody(
        [
            html.Div(
                # id='onelinebody',
                children=[
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test'),
                    html.P('Test')
                ]

            )
        ]
    )
)


first_card = dbc.Card(
    dbc.CardBody(
        [
        html.Div(id='corp-eng-control-tabs', className='control-tabs', children=[
            dcc.Tabs(
                id='corp-eng-tabs', value='what-is',
                children=[
                    dcc.Tab(
                        label='Home',
                        value='what-is',
                        children=html.Div(className='control-tab', children=[
                            html.P(
                                "The adjacent tabs contain tools to be used by the Corporate Engineering Team",
                                style={"font-weight": "bold"}
                            ),
                            html.Br(),
                            html.P(
                                """
                                Oneline: Excel formatted Aries Oneline
                                """
                            ),
                            html.P(
                                """
                                RMR: Reserve Management Record Input/Sign Off Sheets
                                """
                            ),
                            html.P(
                                """
                                Wiki: Information to access specific templates/files/queries/etc.
                                """
                            ),
                        ])
                    ),
                    dcc.Tab(
                        label='Oneline',
                        value='oneline-select',
                        children=html.Div(className='control-tab', children=[
                            html.Label('Enter Scenario', style={"font-weight": "bold"}),
                            dbc.Input(id='idscenario', type='text'),
                            html.Br(),
                            html.Br(),
                            html.Label('Save as', style={"font-weight": "bold"}),
                            dbc.Input(id='idsaveas', type='text'),
                            html.Br(),
                            html.Br(),
                            html.Br(),
                            html.Button("Create Oneline Report", id='onelinebtn', n_clicks=0),
                            html.Br(),
                            html.Br(),
                            html.Button("Download Oneline Report", id='downloadbtn', n_clicks=0),
                            html.Div(id='create_oneline', children=None)
                        ])
                    ),
                    dcc.Tab(
                        label='RMR',
                        value='rmr-select',
                        children=html.Div(className='control-tab', children=[
                            html.Div(className='app-controls-block', children=[
                                html.Div(
                                    className='fullwidth-app-controls-name',
                                    children="Select RMR Phase"
                                ),
                                dcc.Dropdown(
                                    id='alignment-dropdown',
                                    options=[
                                        {
                                            'label': 'RMR Email',
                                            'value': 'rmremail'
                                        },
                                        {
                                            'label': 'RMR Check Status',
                                            'value': 'rmrcheckstatus'
                                        },
                                        {
                                            'label': 'RMR Finalize',
                                            'value': 'rmrfinalize'
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
                                    children="Corporate Engineering Wiki"
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
                id='onelinebody',
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
                                 value=''),
                    html.Br(),
                    html.Label('Select Effective Date, Year', style={"font-weight": "bold"}),
                    dcc.Dropdown(id='idyear',
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
                                 value='')
                ]
            )
        ]
    )
)


def layout():
    return html.Div(id='corp-eng-body', className='app-body', children=[
        dbc.CardGroup(
            [
                dbc.Row(
                    [
                        dbc.Col(first_card, style={'display': 'inline-block'}),
                        dbc.Col(second_card, style={'display': 'inline-block'})
                    ]
                )
            ]
        )
    ])


def callbacks(_app):
    @app.callback(
        Output("create_oneline", "children"),
        [Input("onelinebtn", "n_clicks")],
        [State('iddatabase', 'value'), State('idaconeline', 'value'), State('idacproperty', 'value'),
         State('idscenario', 'value'),
         State('ideffmonth', 'value'), State('ideffyear', 'value'), State('iduser', 'value'),
         State('idsaveas', 'value')]
    )
    def buttons(onelinebtn, database, aconeline, acproperty, scenario, month, year, user, saveas):
        if onelinebtn:
            filepath = Oneline.oneline_report(database, aconeline, acproperty, scenario, month, year, user, saveas)
            a = "File is saved in the file path below. "
            return a, filepath
    @_app.callback(
        Output('onelinebody', 'style'),
        [Input('corp-eng-tabs', 'value')]
    )
    def show_hide_custom_colorscheme(oneline_data):
        return {'display': 'block' if oneline_data == 'oneline-select' else 'none'}

    # @app.callback(
    #     Output("create_oneline", "children"),
    #     [Input("onelinebtn", "n_clicks")],
    #     [State('iddatabase', 'value'), State('idaconeline', 'value'), State('idacproperty', 'value'),
    #      State('idscenario', 'value'),
    #      State('ideffmonth', 'value'), State('ideffyear', 'value'), State('iduser', 'value'),
    #      State('idsaveas', 'value')]
    # )
    # def buttons(onelinebtn, database, aconeline, acproperty, scenario, month, year, user, saveas):
    #     if onelinebtn:
    #         filepath = Oneline.oneline_report(database, aconeline, acproperty, scenario, month, year, user, saveas)
    #         a = "File is saved in the file path below. "
    #         return a, filepath


app = run_standalone_app(layout, callbacks, header_colors, __file__)
server = app.server

if __name__ == '__main__':
    app.run_server(debug=True, port=8050)
