from Queries import *
from Util import *
import xlwings as xw


def oneline_report(database, aconeline, acproperty, scenario, effmonth, effyear, user, filename):
    config = read_yaml("//enc-azfs01/AriesData/CORP_ENG/10 Tools/03 Oneline/config.yaml")
    access_conn = open_db_connection_trusted(config['connection_string'][database])
    templatepath = '//enc-azfs01/AriesData/CORP_ENG/10 Tools/03 Oneline/Oneline_Template.xlsx'
    xw.App.visible = False
    path = '//enc-azfs01/AriesData/CORP_ENG/10 Tools/03 Oneline/Outputs/'+user+'_'
    workbook = xw.Book(templatepath)

    # Variables
    oneline = aconeline
    effdate = effmonth + '-' + effyear
    rescat = 'RSV_CAT_RSV'
    lease = 'WELL_NAME'
    api = 'ID_API_NUM'
    wi = 'M31'
    nri = 'M30'
    grossoil = 'C370'
    grossgas = 'C371'
    grossngl = 'C374'
    netoil = 'C815'
    netgas = 'C816'
    netngl = 'C819'
    netrev = 'C861'
    sevtax = 'C887'
    advaltax = 'C1064'
    netopex = 'C1062'
    netcapex = 'C1183'
    netaband = 'C1094'
    netsalv = 'C1092'
    pv1rate = '0'
    pv2rate = '.05'
    pv3rate = '.10'
    pv4rate = '.15'
    pv5rate = '.20'
    PV1 = 'B1'
    PV2 = 'B2'
    PV3 = 'B4'
    PV4 = 'B6'
    PV5 = 'B7'

    # SQL Queries
    onelinedatadf = OnelineSQL(access_conn, aconeline, acproperty, rescat, lease, api, wi, nri, grossoil, grossgas,
                               grossngl, netoil, netgas, netngl, netrev, sevtax, advaltax, netopex, netcapex, netaband,
                               netsalv, PV1, PV2, PV3, PV4, PV5, scenario)

    sumdf = onelinedatadf.sum()
    rescatlist = onelinedatadf[rescat].unique().tolist()

    outputdict = {}
    firstrowint = 8
    for i in enumerate(rescatlist):
        df = onelinedatadf[onelinedatadf[rescat] == i[1]]
        workbook.sheets["Oneline"].range("A" + str(firstrowint)).options(index=False, header=False).value = df
        firstrowint = firstrowint+df.shape[0]
        sums = df.sum()

        workbook.sheets["Oneline"].range("A" + str(firstrowint)).options(index=False, header=False).value = i[1] + ' Total'
        workbook.sheets["Oneline"].range("A" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("G" + str(firstrowint)).options(index=False, header=False).value = sums[grossoil]
        workbook.sheets["Oneline"].range("G" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("H" + str(firstrowint)).options(index=False, header=False).value = sums[grossgas]
        workbook.sheets["Oneline"].range("H" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("I" + str(firstrowint)).options(index=False, header=False).value = sums[grossngl]
        workbook.sheets["Oneline"].range("I" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("J" + str(firstrowint)).options(index=False, header=False).value = sums[netoil]
        workbook.sheets["Oneline"].range("J" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("K" + str(firstrowint)).options(index=False, header=False).value = sums[netgas]
        workbook.sheets["Oneline"].range("K" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("L" + str(firstrowint)).options(index=False, header=False).value = sums[netngl]
        workbook.sheets["Oneline"].range("L" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("M" + str(firstrowint)).options(index=False, header=False).value = sums[netrev]
        workbook.sheets["Oneline"].range("M" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("N" + str(firstrowint)).options(index=False, header=False).value = sums['Taxes']
        workbook.sheets["Oneline"].range("N" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("O" + str(firstrowint)).options(index=False, header=False).value = sums[netopex]
        workbook.sheets["Oneline"].range("O" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("P" + str(firstrowint)).options(index=False, header=False).value = sums[netcapex]
        workbook.sheets["Oneline"].range("P" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("Q" + str(firstrowint)).options(index=False, header=False).value = sums['P&A']
        workbook.sheets["Oneline"].range("Q" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("R" + str(firstrowint)).options(index=False, header=False).value = sums[PV1]
        workbook.sheets["Oneline"].range("R" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("S" + str(firstrowint)).options(index=False, header=False).value = sums[PV2]
        workbook.sheets["Oneline"].range("S" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("T" + str(firstrowint)).options(index=False, header=False).value = sums[PV3]
        workbook.sheets["Oneline"].range("T" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("U" + str(firstrowint)).options(index=False, header=False).value = sums[PV4]
        workbook.sheets["Oneline"].range("U" + str(firstrowint)).api.Font.Bold = True
        workbook.sheets["Oneline"].range("V" + str(firstrowint)).options(index=False, header=False).value = sums[PV5]
        workbook.sheets["Oneline"].range("V" + str(firstrowint)).api.Font.Bold = True
        firstrowint = firstrowint + 2

    # Grand Totals
    workbook.sheets["Oneline"].range("A" + str(firstrowint)).options(index=False, header=False).value = 'Grand Total'
    workbook.sheets["Oneline"].range("A" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("G" + str(firstrowint)).options(index=False, header=False).value = sumdf[grossoil]
    workbook.sheets["Oneline"].range("G" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("H" + str(firstrowint)).options(index=False, header=False).value = sumdf[grossgas]
    workbook.sheets["Oneline"].range("H" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("I" + str(firstrowint)).options(index=False, header=False).value = sumdf[grossngl]
    workbook.sheets["Oneline"].range("I" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("J" + str(firstrowint)).options(index=False, header=False).value = sumdf[netoil]
    workbook.sheets["Oneline"].range("J" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("K" + str(firstrowint)).options(index=False, header=False).value = sumdf[netgas]
    workbook.sheets["Oneline"].range("K" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("L" + str(firstrowint)).options(index=False, header=False).value = sumdf[netngl]
    workbook.sheets["Oneline"].range("L" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("M" + str(firstrowint)).options(index=False, header=False).value = sumdf[netrev]
    workbook.sheets["Oneline"].range("M" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("N" + str(firstrowint)).options(index=False, header=False).value = sumdf['Taxes']
    workbook.sheets["Oneline"].range("N" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("O" + str(firstrowint)).options(index=False, header=False).value = sumdf[netopex]
    workbook.sheets["Oneline"].range("O" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("P" + str(firstrowint)).options(index=False, header=False).value = sumdf[netcapex]
    workbook.sheets["Oneline"].range("P" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("Q" + str(firstrowint)).options(index=False, header=False).value = sumdf['P&A']
    workbook.sheets["Oneline"].range("Q" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("R" + str(firstrowint)).options(index=False, header=False).value = sumdf[PV1]
    workbook.sheets["Oneline"].range("R" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("S" + str(firstrowint)).options(index=False, header=False).value = sumdf[PV2]
    workbook.sheets["Oneline"].range("S" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("T" + str(firstrowint)).options(index=False, header=False).value = sumdf[PV3]
    workbook.sheets["Oneline"].range("T" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("U" + str(firstrowint)).options(index=False, header=False).value = sumdf[PV4]
    workbook.sheets["Oneline"].range("U" + str(firstrowint)).api.Font.Bold = True
    workbook.sheets["Oneline"].range("V" + str(firstrowint)).options(index=False, header=False).value = sumdf[PV5]
    workbook.sheets["Oneline"].range("V" + str(firstrowint)).api.Font.Bold = True

    # Write to workbook
    filename = filename+'.xlsx'
    write = path + filename
    workbook.sheets["Oneline"].range("A2").options(index=False).value = 'EFFECTIVE DATE: ' + effdate
    workbook.sheets["Oneline"].range("R4").options(index=False).value = pv1rate
    workbook.sheets["Oneline"].range("S4").options(index=False).value = pv2rate
    workbook.sheets["Oneline"].range("T4").options(index=False).value = pv3rate
    workbook.sheets["Oneline"].range("U4").options(index=False).value = pv4rate
    workbook.sheets["Oneline"].range("V4").options(index=False).value = pv5rate
    workbook.save(path=write)
    workbook.close()
    print('Oneline file created')
    return write


# oneline_report('ARIES_RSV', 'AC_ONELINE', 'AC_PROPERTY', 'EAP_Q321_SEC', '10', '2021', 'jddearagao', 'Testing')
