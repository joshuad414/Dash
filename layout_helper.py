import base64
import os

import dash
from dash import dcc
from dash import html


def run_standalone_app(
        layout,
        callbacks,
        header_colors,
        filename
):
    """Run demo app (tests/dashbio_demos/*/app_old.py) as standalone app."""
    app = dash.Dash(__name__)
    app.scripts.config.serve_locally = True
    # Handle callback to component with id "fullband-switch"
    app.config['suppress_callback_exceptions'] = True

    # Get all information from filename
    app_name = 'Encino Corporate Engineering App'
    app_title = app_name

    # Assign layout
    app.layout = app_page_layout(
        page_layout=layout(),
        app_title=app_title,
        app_name=app_name,
        standalone=True,
        **header_colors()
    )

    # Register all callbacks
    callbacks(app)

    # return app object
    return app


def app_page_layout(page_layout,
                    app_title="Encino Corporate Engineering App",
                    app_name="Encino",
                    light_logo=False,
                    standalone=False,
                    bg_color="#3e4827",
                    font_color="#F3F6FA"):
    return html.Div(
        id='main_page',
        children=[
            dcc.Location(id='url', refresh=False),
            html.Div(
                id='app-page-header',
                children=[
                    html.A(
                        id='dashbio-logo', children=[
                            html.Img(
                                src='data:image/png;base64,{}'.format(
                                    base64.b64encode(
                                        open(
                                            './assets/Encino-logo.png', 'rb'
                                        ).read()
                                    ).decode()
                                )
                            )],
                        href="https://www.encinoenergy.com/"
                    ),
                    html.H2(
                        app_title
                    ),

                    html.A(
                        id='gh-link',
                        children=[
                            'Go to Corp Eng Site'
                        ],
                        href="https://encinoenergy.sharepoint.com/sites/Intranet/CorporateEngineering-Private/Shared%20Documents/Forms/AllItems.aspx",
                        style={'color': 'white' if light_logo else 'white',
                               'border': 'solid 1px white' if light_logo else 'solid 1px white'}
                    ),

                    html.Img(
                        src='data:image/png;base64,{}'.format(
                            base64.b64encode(
                                open(
                                    './assets/Sharepoint.png'.format(
                                        'Light-' if light_logo else ''
                                    ),
                                    'rb'
                                ).read()
                            ).decode()
                        )
                    )
                ],
                style={
                    'background': bg_color,
                    'color': font_color,
                }
            ),
            html.Div(
                id='app-page-content',
                children=page_layout
            )
        ],
    )
