import os
import zipfile
import win32com.client
import xlwings as xw
import glob
import pandas as pd
import numpy as np
import subprocess
from tabula import read_pdf
import urllib.request



def move_files(year, quarter):
    rmrpath = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/'
    dirpath = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves'

    os.replace(rmrpath+"D&C Reserves Input Template.xlsm",
               dirpath+"/02 CAPEX/Drilling Completion CAPEX " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"Facilities Capex Reserves Input Template.xlsm",
               dirpath+"/02 CAPEX/Facilities CAPEX " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"PDP Ownership Reserves Input Template.xlsm",
               dirpath+"/04 Ownership/PDP Ownership " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"PDNP Ownership Reserves Input Template.xlsm",
               dirpath+"/04 Ownership/PDNP Ownership " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"POD Reserves Input Template.xlsm",
               dirpath+"/04 Ownership/POD " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"Pricing Reserves Input Template.xlsm",
               dirpath+"/03 Pricing/Pricing " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"OP Schedule Reserves Input Template.xlsm",
               dirpath+"/05 Schedule/OP Schedule " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"Undev Shrink Yield Reserves Input Template.xlsm",
               dirpath+"/07 Shrink Yield/Undev Shrink Yield " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"Undeveloped Forecasts Reserves Input Template.xlsm",
               dirpath+"/06 Forecast/03 PUD/Undeveloped Forecasts " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"BTU Reserves Input Template.xlsm",
               dirpath+"/03 Pricing/BTU " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"Deducts Diffs Reserves Input Template.xlsm",
               dirpath+"/01 LOE/Deducts Diffs " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"OPEX Reserves Input Template.xlsm",
               dirpath+"/01 LOE/OPEX " + quarter + " " + year +".xlsm")
    os.replace(rmrpath+"PDP Shrink Yield Reserves Input Template.xlsm",
               dirpath+"/07 Shrink Yield/PDP Shrink Yield " + quarter + " " + year +".xlsm")
    return 'Files Moved'


def zip_files(year, quarter):
    dirpath = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/'
    new_file = dirpath + '.zip'
    zip =zipfile.ZipFile(new_file, 'w', zipfile.ZIP_DEFLATED)
    for dirpath, dir_names, files in os.walk(dirpath):
        f_path = dirpath.replace(dirpath, '')
        f_path = f_path and f_path + os.sep
        if dirpath == '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/01 Templates':
            break
        else:
            for file in files:
                zip.write(os.path.join(dirpath, file), f_path + file)
    zip.close()

    os.replace('//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/.zip',
               '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/Final_RMR_Data.zip')
    return 'Files zipped'


def email_data(mailto, subject, body, year, quarter):
    rmr_form = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/Reserves Management Form.pdf'
    zip_file = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/Final_RMR_Data.zip'
    get_RMR_Form(year, quarter)
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailto
    mail.Subject = subject
    email_body = '<p style="font-family:palatino linotype;">'+body+'</p>'
    mail.HTMLBody = email_body+'<p style="font-family:palatino linotype;">Thank you,' \
                    '<br style="font-family:palatino linotype; color:rgb(62,72,39)"><b>Corporate Engineering</b></p>' \
                    '<p style="font-family:palatino linotype;">Encino Energy<br>5847 San Felipe Street, Suite 400<br>Houston, TX 77057<br>www.encinoenergy.com</p>'
    mail.Attachments.Add(rmr_form)
    mail.Attachments.Add(zip_file)
    mail.Send()
    os.remove(rmr_form)
    os.remove(zip_file)
    return 'Email sent'


def get_RMR_Form(year, quarter):
    in_file = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/Reserves Management Form.docx'
    out_file = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/Reserves Management Form.pdf'
    wdFormatPDF = 17

    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


def merge_files(year, quarter):
    combine_pricing_data(year, quarter)
    os.remove('//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/Oil Pricing Reserves Input Template.xlsm')
    os.remove('//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/Gas Pricing Reserves Input Template.xlsm')

    path = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/'
    extension = 'xlsm'
    excel_files = glob.glob(path + '*.{}'.format(extension))
    xw.App.visible = False
    combined_wb = xw.Book()

    for excel_file in excel_files:
        wb = xw.Book(excel_file)
        print(excel_file)
        for sheet in wb.sheets:
            sheet.copy(after=combined_wb.sheets[0])
        wb.close()

    combined_wb.sheets[0].delete()
    combined_wb.save(path+quarter+' '+year+' Reserve Inputs.xlsx')
    if len(combined_wb.app.books) == 1:
        combined_wb.app.quit()
    else:
        combined_wb.close()

    return 'Pricing Files Combined'


def download_sec_file(download_url, year, quarter):
    response = urllib.request.urlopen(download_url)
    price_file = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/03 Pricing/' + quarter + ' ' + year +' SEC Pricing.pdf'
    file = open(price_file, 'wb')
    file.write(response.read())
    file.close()


def aries_csv(year, quarter):
    pdp_own = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/04 OWnership/PDP Ownership ' + quarter + ' ' + year +'.xlsm'
    df_pdp_own = pd.read_excel(pdp_own)
    df_pdp_own = df_pdp_own.iloc[1:, :]
    df_pdp_own = df_pdp_own[['Unnamed: 1', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']]
    df_pdp_own = df_pdp_own.rename({'Unnamed: 1': 'QUORUM_ID', 'Unnamed: 6': 'WI', 'Unnamed: 7': 'NRI', 'Unnamed: 8': 'GMI'}, axis=1)
    df_pdp_own.index.names = ['idx']
    df_pdp_own['ReserveQuarter'] = quarter + ' ' + year
    df_pdp_own.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/PDP Ownership.csv')

    pdnp_own = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/04 Ownership/PDNP Ownership '+ quarter + ' ' + year +'.xlsm'
    df_pdnp_own = pd.read_excel(pdnp_own)
    df_pdnp_own = df_pdnp_own.iloc[2:, :]
    df_pdnp_own = df_pdnp_own[['PDNP Ownership', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']]
    df_pdnp_own = df_pdnp_own.rename({'PDNP Ownership': 'PROPNUM', 'Unnamed: 7': 'WI', 'Unnamed: 8': 'NRI', 'Unnamed: 9': 'GMI'}, axis=1)
    df_pdnp_own.index.names = ['idx']
    df_pdnp_own['ReserveQuarter'] = quarter + ' ' + year
    df_pdnp_own.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/PDNP Ownership.csv')

    pdp_sy = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/07 Shrink Yield/PDP Shrink Yield ' + quarter + ' ' + year +'.xlsm'
    df_pdp_sy = pd.read_excel(pdp_sy)
    df_pdp_sy = df_pdp_sy.iloc[1:, :]
    df_pdp_sy = df_pdp_sy[['PDP Shrink/NGL Yields', 'Unnamed: 4', 'Unnamed: 5']]
    df_pdp_sy = df_pdp_sy.rename({'PDP Shrink/NGL Yields': 'PROPNUM', 'Unnamed: 4': 'SHRINK', 'Unnamed: 5': 'YIELD'}, axis=1)
    df_pdp_sy.index.names = ['idx']
    df_pdp_sy['ReserveQuarter'] = quarter + ' ' + year
    df_pdp_sy.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/PDP Shrink Yield.csv')

    undev_sy = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/07 Shrink Yield/Undev Shrink Yield ' + quarter + ' ' + year +'.xlsm'
    df_undev_sy = pd.read_excel(undev_sy)
    df_undev_sy = df_undev_sy.iloc[1:, :]
    df_undev_sy = df_undev_sy[['Undeveloped Shrink/NGL Yields', 'Unnamed: 4', 'Unnamed: 5']]
    df_undev_sy = df_undev_sy.rename({'Undeveloped Shrink/NGL Yields': 'PROPNUM', 'Unnamed: 4': 'SHRINK', 'Unnamed: 5': 'YIELD'}, axis=1)
    df_undev_sy.index.names = ['idx']
    df_undev_sy['ReserveQuarter'] = quarter + ' ' + year
    df_undev_sy.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/Undev Shrink Yield.csv')

    df_lookup = pd.DataFrame(columns=['NAME', 'LINETYPE', 'SEQUENCE', 'VAR0', 'VAR1', 'VAR2', 'VAR3', 'VAR4', 'VAR5', 'VAR6', 'VAR7', 'VAR8', 'VAR9', 'VAR10'])
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '0', 'VAR0': 'TEXT'
                                  ,'VAR1': 'GAS GATHERING'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '1',
                                  'VAR0': 'S/196', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/196'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '2', 'VAR0': 'TEXT'
                                  ,'VAR1': 'GAS COMPRESSION'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '3',
                                  'VAR0': 'S/196', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/196'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '4', 'VAR0': 'TEXT'
                                  ,'VAR1': 'OIL PROCESS & TRANS'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '5',
                                  'VAR0': 'S/195', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/195'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '6', 'VAR0': 'TEXT'
                                  ,'VAR1': 'NGL GATHERING'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '7',
                                  'VAR0': 'S/199', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/199'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '8', 'VAR0': 'TEXT'
                                  ,'VAR1': 'NGL TRANSPORT'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '9',
                                  'VAR0': 'S/199', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/199'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '10', 'VAR0': 'TEXT'
                                  ,'VAR1': 'NGL PROCESS'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '11',
                                  'VAR0': 'S/199', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/199'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '12', 'VAR0': 'TEXT'
                                  ,'VAR1': 'NGL MARKETING'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '13',
                                  'VAR0': 'S/199', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/199'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '14', 'VAR0': 'TEXT'
                                  ,'VAR1': 'NGL OTHER'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '15',
                                  'VAR0': 'S/199', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE',
                                  'VAR6': 'PLUS', 'VAR7': 'S/199'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '0',
                                  'VAR0': 'OPERATED', 'VAR1': 'WET_DRY', 'VAR2': 'GAS_GATH', 'VAR3': 'GAS_COMP_OTH', 'VAR4': 'OPL_P_T', 'VAR5': 'NGL_GATH',
                                  'VAR6': 'NGL_TRANS', 'VAR7': 'NGL_PROCESS', 'VAR8': 'NGL_MARK', 'VAR9': 'NGL_OTH'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'DEDUCTS_'+ quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '1',
                                  'VAR0': 'M', 'VAR1': 'M', 'VAR2': 'C', 'VAR3': 'C', 'VAR4': 'C', 'VAR5': 'N',
                                  'VAR6': 'N', 'VAR7': 'N', 'VAR8': 'N', 'VAR9': 'N'}, ignore_index=True)
    ded = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/01 LOE/Deducts Diffs ' + quarter + ' ' + year +'.xlsm'
    ded = pd.read_excel(ded)
    ded = ded.iloc[1:, :]
    ded = ded[['Deducts & Differentials', 'Unnamed: 1', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11',
                           'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15']]
    ded = ded.rename({'Deducts & Differentials': 'VAR0', 'Unnamed: 1': 'VAR1','Unnamed: 8': 'VAR2', 'Unnamed: 9': 'VAR3', 'Unnamed: 10': 'VAR4', 'Unnamed: 11': 'VAR5',
                           'Unnamed: 12': 'VAR6', 'Unnamed: 13': 'VAR7', 'Unnamed: 14': 'VAR8', 'Unnamed: 15': 'VAR9'}, axis=1)
    convert_dict = {'VAR2': float, 'VAR3': float, 'VAR4': float, 'VAR5': float, 'VAR6': float, 'VAR7': float, 'VAR8': float, 'VAR9': float}
    ded = ded.astype(convert_dict)
    ded = ded.round({'VAR2':4, 'VAR3':4, 'VAR4':4, 'VAR5':4, 'VAR6':4, 'VAR7':4, 'VAR8':4, 'VAR9':4})
    ded['LINETYPE'] = 3
    ded['SEQUENCE'] = ded.index-1
    ded['NAME'] = 'DEDUCTS_'+quarter+ year[-2:]

    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '0'
                                  , 'VAR0': 'OPC/T', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '1'
                                  , 'VAR0': 'OPC/GAS', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '2'
                                  , 'VAR0': 'OPC/WTR', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '3'
                                  , 'VAR0': 'OPC/GAS', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '4'
                                  , 'VAR0': 'OPC/OIL', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '5'
                                  , 'VAR0': 'ABAN', 'VAR1': '220', 'VAR2': 'X', 'VAR3': 'M$', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '6'
                                  , 'VAR0': 'SALV', 'VAR1': '20', 'VAR2': 'X', 'VAR3': 'M$', 'VAR4': 'TO', 'VAR5': 'LIFE'
                                  , 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '0',
                                  'VAR0': 'OPERATED', 'VAR1': 'PAHSE_WINDOW', 'VAR2': 'FIXED', 'VAR3': 'VAR_GAS', 'VAR4': 'VAR_WTR', 'VAR5': 'VAR_OTH',
                                  'VAR6': 'VAR_OIL'}, ignore_index=True)
    df_lookup = df_lookup.append({'NAME': 'OPEX_'+ quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '1',
                                  'VAR0': 'M', 'VAR1': 'M', 'VAR2': 'C', 'VAR3': 'C', 'VAR4': 'C', 'VAR5': 'C',
                                  'VAR6': 'C'}, ignore_index=True)
    opex = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/01 LOE/OPEX ' + quarter + ' ' + year +'.xlsm'
    opex = pd.read_excel(opex)
    opex = opex.iloc[1:, :]
    opex = opex[['Operating Expenses', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
                           'Unnamed: 5', 'Unnamed: 6']]
    opex = opex.rename({'Operating Expenses': 'VAR0', 'Unnamed: 1': 'VAR1', 'Unnamed: 2': 'VAR2',
                        'Unnamed: 3': 'VAR3', 'Unnamed: 4': 'VAR4', 'Unnamed: 5': 'VAR5', 'Unnamed: 6': 'VAR6', }, axis=1)
    convert_dict = {'VAR2': float, 'VAR3': float, 'VAR4': float, 'VAR5': float, 'VAR6': float}
    opex = opex.astype(convert_dict)
    opex = opex.round({'VAR2': 4, 'VAR3': 4, 'VAR4': 4, 'VAR5': 4, 'VAR6': 4})
    opex['LINETYPE'] = 3
    opex['SEQUENCE'] = opex.index-1
    opex['NAME'] = 'OPEX_'+quarter + year[-2:]
    opex.round({'VAR2': 4, 'VAR3': 4, 'VAR4': 4, 'VAR5': 4, 'VAR6': 4})

    shrink_btu_data = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/03 Pricing/BTU ' + quarter + ' ' + year + '.xlsm'
    shrink_btu_data = pd.read_excel(shrink_btu_data)
    shrink_btu_data = shrink_btu_data.iloc[1:, :]
    shrink_btu_data = shrink_btu_data[['Unnamed: 1']]
    shrink_btu = pd.DataFrame(columns=['NAME', 'LINETYPE', 'SEQUENCE', 'VAR0', 'VAR1', 'VAR2', 'VAR3'])
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_'+ quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '0', 'VAR0': 'BTU', 'VAR1': '?'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '1', 'VAR0': 'SHRINK', 'VAR1': '?'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '0', 'VAR0': 'OPERATED', 'VAR1': 'WET_DRY', 'VAR2': 'BTU', 'VAR3': 'SHRINK'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '1', 'VAR0': 'M', 'VAR1': 'M', 'VAR2': 'C', 'VAR3': 'L'}, ignore_index=True)

    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '0', 'VAR0': 'OP', 'VAR1': 'WET', 'VAR2': shrink_btu_data['Unnamed: 1'].values[0], 'VAR3': '@M.SHRINK_GAS_RSV'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '1', 'VAR0': 'OP','VAR1': 'DRY', 'VAR2': shrink_btu_data['Unnamed: 1'].values[1], 'VAR3': '@M.SHRINK_GAS_RSV'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '2', 'VAR0': 'NON-OP', 'VAR1': 'WET', 'VAR2': shrink_btu_data['Unnamed: 1'].values[0], 'VAR3': '@M.SHRINK_GAS_RSV'}, ignore_index=True)
    shrink_btu = shrink_btu.append({'NAME': 'SHRINK_BTU_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '3', 'VAR0': 'NON-OP', 'VAR1': 'DRY', 'VAR2': shrink_btu_data['Unnamed: 1'].values[1], 'VAR3': '@M.SHRINK_GAS_RSV'}, ignore_index=True)

    taxes = pd.DataFrame(columns=['NAME', 'LINETYPE', 'SEQUENCE', 'VAR0', 'VAR1', 'VAR2', 'VAR3', 'VAR4', 'VAR5', 'VAR6', 'VAR7'])
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '0', 'VAR0': 'STX/OIL', 'VAR1': '0.20', 'VAR2': 'X', 'VAR3': '$/B', 'VAR4': 'TO', 'VAR5': 'LIFE', 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '1', 'VAR0': 'STX/GAS', 'VAR1': '0.03', 'VAR2': 'X', 'VAR3': '$/M', 'VAR4': 'TO', 'VAR5': 'LIFE', 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '0', 'SEQUENCE': '2', 'VAR0': 'ATX', 'VAR1': '?', 'VAR2': 'X', 'VAR3': '%M', 'VAR4': 'TO', 'VAR5': 'LIFE', 'VAR6': 'PC', 'VAR7': '0'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '0', 'VAR0': 'WET_DRY', 'VAR1': 'ATX'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '1', 'SEQUENCE': '1', 'VAR0': 'M', 'VAR1': 'C'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '0', 'VAR0': 'WET', 'VAR1': '0.62'}, ignore_index=True)
    taxes = taxes.append({'NAME': 'TAXES_' + quarter + year[-2:], 'LINETYPE': '3', 'SEQUENCE': '1', 'VAR0': 'DRY', 'VAR1': '1.04'}, ignore_index=True)

    frames = [df_lookup, opex, shrink_btu, taxes]
    df_lookup = pd.concat(frames)
    df_lookup.index.names = ['idx']
    df_lookup['ReserveQuarter'] = quarter + ' ' + year
    df_lookup.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/ARLOOKUP.csv')

    pricing = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/03 Pricing/Pricing ' + quarter + ' ' + year + '.xlsm'
    pricing = pd.read_excel(pricing)
    pricing = pricing.iloc[1:, :]
    df_sidefile = pd.DataFrame(columns=['FILENAME', 'SECTION', 'SEQUENCE', 'KEYWORD', 'EXPRESSION'])
    date = '01/2021'
    count = 2

    # oil pricing
    for i in range(0, len(pricing)):
        if np.remainder(i, 4) != 0 or i == 0:
            date += ' ' + str(np.round(pricing['Unnamed: 2'].values[i], 2))
            keyword = '"'
            if i == 3:
                keyword = 'PRI/OIL'
        if np.remainder(i, 4) == 0 and i != 0:
            date += ' #/M'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': keyword,
                 'EXPRESSION': date}, ignore_index=True)
            date = 'X ' + str(np.round(pricing['Unnamed: 2'].values[i], 2))
            count += 2
        if i >= 61:
            x = pricing.iloc[:, 0][i]
            date = str(np.round(pricing['Unnamed: 2'].values[i], 2)) + ' X $/B 01/' + x[-4:] + ' AD PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)
            count += 2
        if i == len(pricing) - 1:
            date = str(np.round(pricing['Unnamed: 2'].values[i], 2)) + ' X $/B TO LIFE PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)
            date = '01/2021'
            count+= 2

    # gas pricing
    for i in range(0, len(pricing)):
        if np.remainder(i, 4) != 0 or i == 0:
            date += ' ' + str(pricing['Unnamed: 1'].values[i])
            keyword = '"'
            if i == 3:
                keyword = 'PRI/GAS'
        if np.remainder(i, 4) == 0 and i != 0:
            date += ' #/M'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': keyword,
                 'EXPRESSION': date}, ignore_index=True)
            date = 'X ' + str(np.round(pricing['Unnamed: 1'].values[i], 3))
            count += 2
        if i >= 61:
            x = pricing.iloc[:, 0][i]
            date = str(np.round(pricing['Unnamed: 1'].values[i], 3)) + ' X $/M 01/' + x[-4:] + ' AD PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)
            count += 2
        if i == len(pricing) - 1:
            date = str(np.round(pricing['Unnamed: 1'].values[i], 3)) + ' X $/M TO LIFE PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)
            date = '01/2021'
            count += 2

    # ngl pricing
    for i in range(0, len(pricing)):
        if np.remainder(i, 4) != 0 or i == 0:
            date += ' ' + str(np.round(pricing['Unnamed: 2'].values[i], 2))
            keyword = '"'
            if i == 3:
                keyword = 'PRI/NGL'
        if np.remainder(i, 4) == 0 and i != 0:
            date += ' #/M'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': keyword,
                 'EXPRESSION': date}, ignore_index=True)
            date = 'X ' + str(np.round(pricing['Unnamed: 2'].values[i], 2))
            count += 2
        if i >= 61:
            x = pricing.iloc[:, 0][i]
            date = str(np.round(pricing['Unnamed: 2'].values[i], 2)) + ' X $/B 01/' + x[-4:] + ' AD PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)
            count += 2
        if i == len(pricing) - 1:
            date = str(np.round(pricing['Unnamed: 2'].values[i], 2)) + ' X $/B TO LIFE PC 0'
            df_sidefile = df_sidefile.append(
                {'FILENAME': 'NYMEX_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': count, 'KEYWORD': '"',
                 'EXPRESSION': date}, ignore_index=True)

    pdf_file = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/03 Pricing/' + quarter + ' ' + year + ' SEC Pricing.pdf'
    df = read_pdf(pdf_file, pages='2')
    sec_gas_price = df[0]['AVG'].values[20]
    sec_oil_price = df[0]['AVG'].values[4]
    df_sidefile = df_sidefile.append({'FILENAME': 'SEC_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': '25', 'KEYWORD': 'PRI/OIL', 'EXPRESSION': sec_oil_price[1:] + ' X $/B TO LIFE PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'SEC_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': '50', 'KEYWORD': 'PRI/NGL', 'EXPRESSION': sec_oil_price[1:] + ' X $/B TO LIFE PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'SEC_' + quarter + year[-2:], 'SECTION': '5', 'SEQUENCE': '75', 'KEYWORD': 'PRI/GAS', 'EXPRESSION': sec_gas_price[1:] + ' X $/M TO LIFE PC 0'}, ignore_index=True)

    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '12', 'KEYWORD': 'CAPITAL', 'EXPRESSION': 'X @M.CAP_PAD_CE G @M.DT_PRE_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '24', 'KEYWORD': 'DRILL', 'EXPRESSION': 'X @M.CAP_DRILL_CE G @M.DT_SPUD_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '36', 'KEYWORD': 'COMPL', 'EXPRESSION': 'X @M.CAP_COMP_CE G @M.DT_COMP_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '48', 'KEYWORD': 'CAPITAL', 'EXPRESSION': 'X @M.CAP_TIL_CE G @M.DT_FIRST_PROD_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '60', 'KEYWORD': 'CAPITAL', 'EXPRESSION': 'X @M.CAP_TBG_CE G @M.DT_TBG_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '72', 'KEYWORD': 'CAPITAL', 'EXPRESSION': 'X @M.CAP_MISC1_CE G @M.DT_FIRST_PROD_CE AD PC 0'}, ignore_index=True)
    df_sidefile = df_sidefile.append({'FILENAME': 'CAPEX_' + quarter + year[-2:], 'SECTION': '8', 'SEQUENCE': '84', 'KEYWORD': 'CAPITAL', 'EXPRESSION': 'X 20 G @M.DT_AL1_CE AD PC 0'}, ignore_index=True)

    df_sidefile.index.names = ['idx']
    df_sidefile['ReserveQuarter'] = quarter + ' ' + year
    df_sidefile.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/AR_SIDEFILE.csv')

    rmrpath = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/'
    df_dc = pd.read_excel(rmrpath+'02 CAPEX/Drilling Completion CAPEX ' + quarter + ' ' + year + '.xlsm')
    df_dc = df_dc.iloc[1:, :]
    df_fac = pd.read_excel(rmrpath+'02 CAPEX/Facilities CAPEX ' + quarter + ' ' + year + '.xlsm')

    wet_fac_capex = df_fac.iloc[1]['Unnamed: 1']
    dry_fac_capex = df_fac.iloc[2]['Unnamed: 1']
    facility_months = df_fac.iloc[3]['Unnamed: 1']

    # Drilling Capex for Wet
    df_d_wet = df_dc[['Drilling Ladder', 'Unnamed: 1']]
    df_d_wet = df_d_wet.rename({'Drilling Ladder': 'Lateral_Length', 'Unnamed: 1': 'drill_wet'}, axis=1)
    df_d_wet['Lateral_Length'] = df_d_wet['Lateral_Length'].astype(str).astype(int)
    df_d_wet['drill_wet'] = df_d_wet['drill_wet'].astype(str).astype(int)
    ll = df_d_wet.Lateral_Length
    capex_d_wet = df_d_wet.drill_wet
    drilling_wet_model = np.polyfit(ll, capex_d_wet, 3)
    df_d_wet['Model'] = ''
    df_d_wet['Model'] = 'y = ' + format(drilling_wet_model[0], '.3g') + 'x^3 + ' + format(
        drilling_wet_model[1], '.3g') + 'x^2 + ' + format(drilling_wet_model[2], '.3g') + 'x + ' + str(np.round(drilling_wet_model[3],2))

    # Drilling Capex for Dry
    df_d_dry = df_dc[['Drilling Ladder', 'Unnamed: 2']]
    df_d_dry = df_d_dry.rename({'Drilling Ladder': 'Lateral_Length', 'Unnamed: 2': 'drill_dry'}, axis=1)
    df_d_dry['Lateral_Length'] = df_d_dry['Lateral_Length'].astype(str).astype(int)
    df_d_dry['drill_dry'] = df_d_dry['drill_dry'].astype(str).astype(int)
    ll = df_d_dry.Lateral_Length
    capex_d_dry = df_d_dry.drill_dry
    drilling_dry_model = np.polyfit(ll, capex_d_dry, 3)
    df_d_dry['Model'] = ''
    df_d_dry['Model'] = 'y =' + format(drilling_dry_model[0], '.3g') + 'x^3 + ' + format(
        drilling_dry_model[1], '.3g') + 'x^2 + ' + format(drilling_dry_model[2], '.3g') + 'x + ' + str(np.round(drilling_dry_model[3],2))

    # Completion Capex for Wet
    df_c_wet = df_dc[['Drilling Ladder', 'Unnamed: 5']]
    df_c_wet = df_c_wet.rename({'Drilling Ladder': 'Lateral_Length', 'Unnamed: 5': 'compl_wet'}, axis=1)
    df_c_wet['Lateral_Length'] = df_c_wet['Lateral_Length'].astype(str).astype(int)
    df_c_wet['compl_wet'] = df_c_wet['compl_wet'].astype(str).astype(int)
    ll = df_c_wet.Lateral_Length
    capex_c_wet = df_c_wet.compl_wet
    completion_wet_model = np.polyfit(ll, capex_c_wet, 3)
    df_c_wet['Model'] = ''
    df_c_wet['Model'] = 'y = ' + format(completion_wet_model[0], '.3g') + 'x^3 + ' + format(
        completion_wet_model[1], '.3g') + 'x^2 + ' + format(completion_wet_model[2], '.3g') + 'x + ' + str(np.round(completion_wet_model[3],2))

    # Completion Capex for Dry
    df_c_dry = df_dc[['Drilling Ladder', 'Unnamed: 6']]
    df_c_dry = df_c_dry.rename({'Drilling Ladder': 'Lateral_Length', 'Unnamed: 6': 'compl_dry'}, axis=1)
    df_c_dry['Lateral_Length'] = df_c_dry['Lateral_Length'].astype(str).astype(int)
    df_c_dry['compl_dry'] = df_c_dry['compl_dry'].astype(str).astype(int)
    ll = df_c_dry.Lateral_Length
    capex_d_dry = df_c_dry.compl_dry
    completion_dry_model = np.polyfit(ll, capex_d_dry, 3)
    df_c_dry['Model'] = ''
    df_c_dry['Model'] = 'y = ' + format(completion_dry_model[0], '.3g') + 'x^3 + ' + format(
        completion_dry_model[1], '.3g') + 'x^2 + ' + format(completion_dry_model[2], '.3g') + 'x + ' + str(np.round(completion_dry_model[3],2))

    # Get Undeveloped Wells
    df_pod = pd.read_excel(rmrpath+'04 Ownership/POD ' + quarter + ' ' + year + '.xlsm')
    df_pod = df_pod.iloc[5:, :]
    df_pod = df_pod[['Unnamed: 0', 'Unnamed: 4', 'Unnamed: 28']]
    df_pod = df_pod.rename({'Unnamed: 0': 'PROPNUM', 'Unnamed: 4': 'Lateral_Length', 'Unnamed: 28': 'Gathering_System'}, axis=1)
    df_pod['Drill_Capex'] = 0
    df_pod['Drill_Model'] = ''
    df_pod['Compl_Capex'] = 0
    df_pod['Compl_Model'] = ''
    df_pod['Facility_Capex'] = 0
    df_pod['TIL_Capex'] = 0

    for i in range(0, len(df_pod)):
        lat_length = df_pod.iloc[i]['Lateral_Length']
        if df_pod.iloc[i]['Gathering_System'] == 'CARDINAL':
            df_pod.at[i+5, 'Drill_Capex'] = (((drilling_wet_model[0]*(np.power(lat_length,3)))\
                                            +(drilling_wet_model[1]*(np.power(lat_length,2)))\
                                            +(drilling_wet_model[2]*lat_length)\
                                            +drilling_wet_model[3])*lat_length)/1000
            df_pod.at[i + 5, 'Drill_Model'] = 'y = '+str(drilling_wet_model[0])+'x^3 + '+str(drilling_wet_model[1])+'x^2 + '+str(drilling_wet_model[2])+'x + '+str(drilling_wet_model[3])
            df_pod.at[i+5, 'Compl_Capex'] = (((completion_wet_model[0]*(np.power(lat_length,3)))\
                                            +(completion_wet_model[1]*(np.power(lat_length,2)))\
                                            +(completion_wet_model[2]*lat_length)\
                                            +completion_wet_model[3])*lat_length)/1000
            df_pod.at[i + 5, 'Compl_Model'] = 'y = '+str(completion_wet_model[0])+'x^3 + '+str(completion_wet_model[1])+'x^2 + '+str(completion_wet_model[2])+'x + '+str(completion_wet_model[3])
            df_pod.at[i + 5, 'Facility_Capex'] = wet_fac_capex
        else:
            df_pod.at[i+5, 'Drill_Capex'] = (((drilling_dry_model[0]*(np.power(lat_length,3)))\
                                            +(drilling_dry_model[1]*(np.power(lat_length,2)))\
                                            +(drilling_dry_model[2]*lat_length)\
                                            +drilling_dry_model[3])*lat_length)/1000
            df_pod.at[i + 5, 'Drill_Model'] = 'y = '+str(drilling_dry_model[0])+'x^3 + '+str(drilling_dry_model[1])+'x^2 + '+str(drilling_dry_model[2])+'x + '+str(drilling_dry_model[3])
            df_pod.at[i+5, 'Compl_Capex'] = (((completion_dry_model[0]*(np.power(lat_length,3)))\
                                            +(completion_dry_model[1]*(np.power(lat_length,2)))\
                                            +(completion_dry_model[2]*lat_length)\
                                            +completion_dry_model[3])*lat_length)/1000
            df_pod.at[i + 5, 'Compl_Model'] = 'y = '+str(completion_dry_model[0])+'x^3 + '+str(completion_dry_model[1])+'x^2 + '+str(completion_dry_model[2])+'x + '+str(completion_dry_model[3])
            df_pod.at[i + 5, 'Facility_Capex'] = dry_fac_capex

    df_pod.index.names = ['idx']
    df_pod['ReserveQuarter'] = quarter + ' ' + year
    df_pod.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/Undev Capex.csv')

    capex = [df_d_wet, df_c_wet, df_d_dry, df_c_dry]
    capex_df = pd.concat(capex)
    capex_df.to_csv('//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/02 Aries Upload/CAPEX.csv')

    return df_d_wet, df_d_dry, df_c_wet, df_c_dry


def save_supporting_doc(year, quarter, uploadedfile, uploader):
    rmrpath = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/'+year+' '+quarter+' Reserves/19 RMR/'
    with open(os.path.join(rmrpath, uploadedfile.name), "wb") as f:
        f.write(uploadedfile.getbuffer())
        #f.write(rmrpath+uploader+uploadedfile)
    return 'File Saved'


def combine_pricing_data(year, quarter):
    pricing = '//enc-azfs01/AriesData/CORP_ENG/10 Tools/07 RMR/03 Templates/Pricing Reserves Input Template.xlsm'
    oil_price = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/Oil Pricing Reserves Input Template.xlsm'
    gas_price = '//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/Gas Pricing Reserves Input Template.xlsm'

    df_oil = pd.read_excel(oil_price)
    df_gas = pd.read_excel(gas_price)

    df_oil = df_oil.iloc[1:, :]
    df_oil = df_oil[['Unnamed: 1']]
    df_oil = df_oil.rename({'Unnamed: 1': 'Oil Pricing'}, axis=1)

    df_gas = df_gas.iloc[1:, :]
    df_gas = df_gas[['Unnamed: 1']]
    df_gas = df_gas.rename({'Unnamed: 1': 'Gas Pricing'}, axis=1)

    xw.App.visible = False
    workbook = xw.Book(pricing)

    for i in range(0, len(df_oil)):
        workbook.sheets["Pricing"].range("C" + str(i + 3)).options(index=False, header=False).value = \
        df_oil['Oil Pricing'].values[i]

    for i in range(0, len(df_gas)):
        workbook.sheets["Pricing"].range("B" + str(i + 3)).options(index=False, header=False).value = \
        df_gas['Gas Pricing'].values[i]

    workbook.save(path='//enc-azfs01/AriesData/CORP_ENG/01 Reserves/01 Quarterly/' + year + ' ' + quarter + ' Reserves/19 RMR/Pricing Reserves Input Template.xlsm')
    workbook.close()


def update_aries_sql():
    exec = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\01 exec\AriesRsvQtrInput.exe'

    pdp_own_file = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\PDP Ownership.csv'
    pdnp_own_file = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\PDNP Ownership.csv'
    pdp_sy_file = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\PDP Shrink Yield.csv'
    undev_sy_file = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\Undev Shrink Yield.csv'
    undev_capex_file = r'L:\CORP_ENG\10 Tools\07 RMR\02 Aries Upload\Undev Capex.csv'

    pdp_own_name = r'PDPOwnership'
    pdnp_own_name = r'PDNPOwnership'
    pdp_sy_name = r'PDPShrinkYield'
    undev_sy_name = r'UndevShrinkYield'
    undev_capex_name = r'UndevCapex'

    subprocess.run([exec, pdp_own_file, pdp_own_name])
    subprocess.run([exec, pdnp_own_file, pdnp_own_name])
    subprocess.run([exec, pdp_sy_file, pdp_sy_name])
    subprocess.run([exec, undev_sy_file, undev_sy_name])
    subprocess.run([exec, undev_capex_file, undev_capex_name])
