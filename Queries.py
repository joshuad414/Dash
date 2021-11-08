# contains the SQL queries, set up so that all data is passed from the Main file and returns a Dataframe
# DB_Call(VARIABLES, QUERY_TO_PERFORM)

# imports
import pandas as pd


def OnelineSQLOrig(access_conn, Oneline, Property, ResCat, Lease, Api, Wi, Nri, GrossOil, GrossGas, GrossNGL, NetOil, NetGas, NetNGL, NetRev, SevTax, AdValTax, NetOpex, NetCapex, NetAband, NetSalv, PV1, PV2, PV3, PV4, PV5, Scenario):
    query = "SELECT P." + ResCat + ", P." + Lease + ", P." + Api + ", O.PROPNUM, O." + Wi + ", O." + Nri\
            + ", O." + GrossOil + ", O." + GrossGas + ", O." + GrossNGL + ", O." + NetOil + ", O." + NetGas + ", O." + NetNGL + ", O." + NetRev + ", (O." + SevTax + "+ O." + AdValTax + ") Taxes, O." + NetOpex + ", O." + NetCapex + ", (O." + NetAband + "- O." + NetSalv\
            + ") [P&A], O." + PV1 + ", O." + PV2 + ", O." + PV3 + ", O." + PV4 + ", O." + PV5\
            + " FROM " + Oneline + " AS O INNER JOIN " + Property + " AS P ON O.PROPNUM = P.PROPNUM WHERE O.SCENARIO = '" + Scenario + "' ORDER BY P." + ResCat + "," + Lease
    oneline = pd.read_sql(query, access_conn)
    return oneline


def OnelineSQL(access_conn, Oneline, Property, ResCat, Lease, Api, Wi, Nri, GrossOil, GrossGas, GrossNGL, NetOil, NetGas, NetNGL, NetRev, SevTax, AdValTax, NetOpex, NetCapex, NetAband, NetSalv, PV1, PV2, PV3, PV4, PV5, Scenario):
    query = "SELECT P." + ResCat + ", P." + Lease + ", P." + Api + ", O.PROPNUM, O." + Wi + ", O." + Nri\
            + ", O." + GrossOil + "/1000 "+ GrossOil + ", O." + GrossGas + "/1000 "+ GrossGas + ", O." + GrossNGL + "/1000 "+ GrossNGL + ", O." + NetOil + "/1000 "+ NetOil + ", O." + NetGas + "/1000 "+ NetGas + ", O." + NetNGL + "/1000 "+ NetNGL + ", O." + NetRev + "/1000 "+ NetRev + ", (O." + SevTax + "/1000+ O." + AdValTax + "/1000) Taxes, O." + NetOpex + "/1000 "+ NetOpex + ", O." + NetCapex + "/1000 "+ NetCapex + ", (O." + NetAband + "/1000- O." + NetSalv\
            + "/1000) [P&A], O." + PV1 + "/1000 "+ PV1 + ", O." + PV2 + "/1000 "+ PV2 + ", O." + PV3 + "/1000 "+ PV3 + ", O." + PV4 + "/1000 "+ PV4 + ", O." + PV5\
            + "/1000 "+ PV5 + " FROM " + Oneline + " AS O INNER JOIN " + Property + " AS P ON O.PROPNUM = P.PROPNUM WHERE O.SCENARIO = '" + Scenario + "' ORDER BY P." + ResCat + "," + Lease
    oneline = pd.read_sql(query, access_conn)
    return oneline


def AriesUpdateSQL(access_conn):
    query = "exec dbo.sp_UpdateQtrInput SELECT PROPNUM FROM AC_PROPERTY"
    x = pd.read_sql(query, access_conn)
    return x
