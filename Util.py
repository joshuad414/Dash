# -*- coding: utf-8 -*-
"""
Created on Tue Sep 24 08:07:16 2019

@author: adrummond
"""

import yaml
import pyodbc, urllib
import pandas as pd
import sqlalchemy


def read_yaml(yaml_path):
    '''
    Reads YAML configuration file.

    Parameters:
        yaml_path (string): path to the yaml config file

    Returns:
        config_params (dict): dictionary based on yaml file
    '''
    with open(yaml_path, 'r') as stream:
        try:
            config_params = yaml.safe_load(stream)
            return (config_params)
        except yaml.YAMLError as exc:
            print(exc)


def execute_query_to_pddf(access_conn, query):
    '''
    Executes sql query with pandas that returns pandas data frame

    Parameters:
        access_conn (DB Connection): connection to the database
        query (string): query to execute

    Returns:
        data (pandas.DataFrame): results of the query
    '''
    data = pd.read_sql(query, access_conn)
    return (data)


def create_sql_engine(db_params):
    '''
    Creates SQLAlchemy Engine

    Parameters:
        db_params (dict): a dictionary with following keys: driver, server, database, uid, pwd

    Returns:
        engine (Engine): SQLAlchemy Engine
    '''
    sql_driver = db_params['driver']
    sql_server = db_params['server']
    sql_db = db_params['database']

    conn_string = urllib.parse.quote_plus("DRIVER=" + sql_driver +
                                          ";SERVER=" + sql_server +
                                          ";DATABASE=" + sql_db +
                                          ";Trusted_Connection=yes")
    # pyodbc

    engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect=%s" % conn_string)
    return (engine)


def execute_query(engine, query):
    '''
    Executes sql query with sqlalchemy

    Parameters:
        engine (Engine): SQLAlchemy Engine
        query (string): a query to execute
    '''
    # with engine.begin() as conn:  # TRANSACTION
    #   conn.execute(query)

    engine.execute(query)


def open_db_connection_linux(db_params):
    '''
    Opens connection to the db from Linux using the DSN and parameters passed in a dict

    Parameters:
        db_params (dict): a dictionary with following keys: dsn, uid, pwd

    Returns:
        access_conn (DB Connection): connection to the database
    '''

    sql_dsn = db_params['dsn']
    sql_user = db_params['uid']
    sql_password = db_params['pwd']

    access_conn = pyodbc.connect("DSN=" + sql_dsn +
                                 ";UID=" + sql_user +
                                 ";PWD=" + sql_password)
    return access_conn


def open_db_connection(db_params):
    '''
    Opens connection to the db using the parameters passed in a dict

    Parameters:
        db_params (dict): a dictionary with following keys: driver, server, database, uid, pwd

    Returns:
        access_conn (DB Connection): connection to the database
    '''
    sql_driver = db_params['driver']
    sql_server = db_params['server']
    sql_db = db_params['database']

    print("\nOpening connection to \"", sql_db, "\"...", sep="")
    if db_params['uid'] is not None:
        sql_user = db_params['uid']
        sql_password = db_params['pwd']

        access_conn = pyodbc.connect("DRIVER=" + sql_driver +
                                     ";SERVER=" + sql_server +
                                     ";DATABASE=" + sql_db +
                                     ";UID=" + sql_user +
                                     ";PWD=" + sql_password)
    else:
        access_conn = pyodbc.connect("DRIVER=" + sql_driver +
                                     ";SERVER=" + sql_server +
                                     ";DATABASE=" + sql_db +
                                     ";Trusted_Connection=yes")
    return access_conn


def date_add_sub(date, delta_val, delta_attr):
    '''
    Adds or substructs given amount of time to/from date

    Parameters:
        date (datetime): a base date
        delta_attr (string): relative time attribute; choices: days, hours, minutes, seconds, microseconds
        delta_val (integer): a value for a given delta attribute

    Returns:
        offset_date (datetime): a relative date calculated based on the given parameters
    '''
    from datetime import timedelta

    delta = eval("timedelta(" + delta_attr + "=" + str(delta_val) + ")")
    offset_date = date + delta
    return offset_date


def parse_date(date_string):
    '''
    Takes date or a sequence of dates in string format and parses it to a datetime format

    Parameters:
        date_string (string): date as a string

    Returns:
        date_datetime (datetime): date in a datetime format
        date_list (list(datetime)): a list of dates in datetime format
    '''
    from dateutil.parser import parse

    if type(date_string) is str:
        date_datetime = parse(date_string)
        return date_datetime
    else:
        date_strings = date_string
        date_list = []
        for date_string in date_strings:
            date_list += [parse(date_string)]
        return date_list

def open_db_connection_trusted(db_params):
    '''
    Opens connection to the db using the parameters passed in a dict

    Parameters:
        db_params (dict): a dictionary with following keys: driver, server, database, uid, pwd

    Returns:
        access_conn (DB Connection): connection to the database
    '''
    sql_driver = db_params['driver']
    sql_server = db_params['server']
    sql_db = db_params['database']

    print("\nOpening connection to \"", sql_db, "\"...", sep="")

    access_conn = pyodbc.connect("DRIVER=" + sql_driver +
                                 ";SERVER=" + sql_server +
                                 ";DATABASE=" + sql_db +
                                 ";Trusted_Connection=yes")
    return access_conn
