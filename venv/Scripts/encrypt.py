import pandas as pd
import numpy as np
from itertools import chain
import os
from io import StringIO


import sys
import pandas as pd
from datetime import timedelta

path = os.path.dirname(os.path.realpath(__file__))
path=path+"//Excel_files"
print(path)

import xlsxwriter

from pandas import ExcelWriter
import csv, xlrd
from datetime import date
from datetime import timedelta


import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
import time


# import user

def error(valid, e):
    # print("Resort Number "+group.Resort.unique()+ " "+group['Room Type'].unique()+ " Has overlapping dates in Main Tab")
    return (e)


def etl(fd, rt, ol, lra, ria, bwi, repeat, qa, ww):
    ####################################################  USER INPUT AUDIT REPORT  #############################################################
    audit = []
    list5 = ['FAIRDATES', 'ROOMTYPE', 'Loading Occupancy', 'LRA or Non-LRA', 'Rate Integrity', 'BWIRateCode']
    if fd == '0':
        audit.append('No fairdates')
    else:
        audit.append('Yes')

    if rt == '0':
        audit.append('STANDARD')
    if rt == '1':
        audit.append('Upgrade')
    if rt == '2':
        audit.append('Both Standard and upgrade')

    if ol == '0':
        audit.append('Single Person')
    if ol == '1':
        audit.append('Double Person')
    if ol == '2':
        audit.append('Both Single and Double')

    if lra == '0':
        audit.append('LRA')
    if lra == '1':
        audit.append('Non-LRA')
    if lra == '2':
        audit.append('Both')
    audit.append(ria)
    audit.append(bwi)

    d = pd.read_csv(path + "//User.csv")
    d = d.drop_duplicates()
    d = d.reset_index(drop=True)

    df_audit2 = d.iloc[1:]

    df_h1 = pd.DataFrame({'H1': 1, 'Corporate ID': d.iloc[1]['Corp Acct#']}, index=[0])

    df_h2 = d.groupby(
        ['GDS Rate Codes', 'Multi Rate Code', 'Sabre', 'Worldspan', 'Amaedus', 'Galileo', 'Web']).size().reset_index(
        name='Freq')

    df_h2 = df_h2.drop(['Freq'], axis=1)
    df_h2['H2'] = 2
    df_h2 = df_h2[['H2', 'GDS Rate Codes', 'Multi Rate Code', 'Sabre', 'Worldspan', 'Amaedus', 'Galileo', 'Web']]

    # d.drop_duplicates(subset=['Resort'], keep='first')
    e = 0
    w = 0
    df_Main = pd.read_excel(path + "//MAIN.xlsx", sheet_name='Sheet1')
    df_vc = pd.read_excel(path + "//Valid_Cancel_Codes.xlsx", sheet_name='Sheet1')
    writer1 = pd.ExcelWriter(path + '//Audit_Rpt.xlsx', engine='xlsxwriter')
    audit_dic = {'Attributes': list5, 'User Input': audit}
    df_audit = pd.DataFrame(audit_dic)
    df_audit.to_excel(writer1, sheet_name='User Inputs', index=False)
    df_audit2.to_excel(writer1, sheet_name='Rate code details', index=False)
    df_h1.to_excel(writer1, sheet_name='Rate code h2', index=False)
    df_h2.to_excel(writer1, sheet_name='Rate code h2', index=False, startrow=len(df_h1) + 1)
    writer1.save()
    ####################################################  USER INPUT  #############################################################
    ###########################################################################ROH############################################
    list_CE = ['Austria', 'Croatia', 'Czech Republic', 'Germany', 'Hungary', 'Luxembourg', 'Slovakia', 'Slovenia',
               'Switzerland']
    CE_ROH_NUM = []
    for i in list_CE:
        yy = df_Main.loc[(df_Main['CountryName'] == i) & (df_Main['RoomDescription'] == 'ROH') & (
                    df_Main['RoomTypeName'] == 'STANDARD')]["CRSHotelID"]
        lo = yy.tolist()
        # yy=df_Main.get_value(df_Main[(df_Main['CountryName'] == i )&  (df_Main['RoomDescription'] == 'ROH' )& (df_Main['RoomTypeName'] == 'STANDARD')].index,'CRSHotelID')
        try:
            CE_ROH_NUM.append(lo[0])
        except:
            pass

    for i in CE_ROH_NUM:
        try:
            df_Main = df_Main.drop(df_Main[(df_Main['CRSHotelID'] == i) & (df_Main['RoomTypeName'] == 'DELUXE')].index,
                                   inplace=False)
        except:
            pass
    df_Main = df_Main.reset_index(drop=True)

    #####################################################################################################################
    if fd == '1':
        df_FairDates = pd.read_excel(path + "//FAIR.xlsx", sheet_name='Sheet1')
    else:
        df_FairDates = pd.read_excel(path + '//WorkingFile_3.xlsx', sheet_name='Sheet1')
    df_Blackout = pd.read_excel(path + '//MAIN.xlsx', sheet_name='Sheet3')
    df_Cancel = pd.read_excel(path + '//MAIN.xlsx', sheet_name='Sheet2')
    today = date.today().strftime('%m/%d/%Y')
    valid = pd.DataFrame()
    if repeat != 0:
        e = 5
        df_Main = df_Main[~df_Main['CRSHotelID'].isin(qa)]
        df_Main = df_Main[~df_Main['CRSHotelID'].isin(ww)]
        if fd == '1':
            df_FairDates = df_FairDates[~df_FairDates['PROPCODE'].isin(qa)]
            df_FairDates = df_FairDates[~df_FairDates['PROPCODE'].isin(ww)]
        else:
            pass
        df_Blackout = df_Blackout[~df_Blackout['CRSHotelID'].isin(qa)]
        df_Cancel = df_Cancel[~df_Cancel['CRSHotelID'].isin(qa)]
        d = d[~d["Resort"].isin(qa)]
        df_Blackout = df_Blackout[~df_Blackout['CRSHotelID'].isin(ww)]
        df_Cancel = df_Cancel[~df_Cancel['CRSHotelID'].isin(ww)]
        d = d[~d["Resort"].isin(ww)]
    # qq=pd.DataFrame({ '1H': 1,'2H':qa})
    # return(error(df_Main,e))

    lm = len(df_Main)
    d = d.iloc[1:]
    d = d[["Corp Acct#", "GDS Rate Codes", "Multi Rate Code", "Resort", "BWI Rate Code", "Map to ROH", "Begin Date",
           "End Date", "Sabre", "Worldspan", "Amaedus", "Galileo", "Web"]]
    d.columns = ["Corp Acct #", "GDS Rate Codes", "Multi Rate Codes", "Resort", "BWI Rate Code", "Map to ROH (Y/N)",
                 "Begin Date", "End Date", "Sabre", "WorldSpan", "Amadeus", "Galileo", "Web"]

    d["BWI Rate Code"] = bwi
    idx = 0
    d.insert(loc=idx, column='2H', value=2)
    if d.empty:
        print("All room types and resorts are not matching or have overlapping dates or date gaps")

    H1 = pd.DataFrame({'1H': 1, 'Corp Acct#': d.iloc[0]['Corp Acct #']}, index=[0])

    def chainer(s):
        return list(chain.from_iterable(s.str.split(';')))

    ####################################ROH############################################
    df_Cancel['Base_Rate'] = 'RACK'
    uz_con = ['Austria', 'Croatia', 'Czech Republic', 'Germany', 'Hungary', 'Luxembourg', 'Slovakia', 'Slovenia',
              'Switzerland', 'Algeria', 'France', 'French Guiana', 'Morocco', 'Reunion', 'Channel Islands', 'England',
              'Gibraltar', 'Scotland', 'Wales']
    for i in uz_con:
        df_Cancel.loc[df_Cancel['CountryName'] == i, 'Base Rate'] = 'UZ'
    list_FE = ['Algeria', 'France', 'French Guiana', 'Morocco', 'Reunion']
    FE_ROH_NUM = []
    for i in list_FE:
        df_Cancel.loc[(df_Cancel['CountryName'] == i) & (df_Cancel['CompBreakfast'] == 'Y') & (
                    df_Cancel['CityTaxIncluded'] == 'N'), 'Base Rate'] = '8A'
        df_Cancel.loc[(df_Cancel['CountryName'] == i) & (df_Cancel['CompBreakfast'] == 'N') & (
                    df_Cancel['CityTaxIncluded'] == 'N'), 'Base Rate'] = 'RACK'

    ################################################### Details of the Resort ##################################################################

    df_Cancel.loc[df_Cancel['CountryName'] == 'United States', 'Base Rate'] = 'RACK'
    df_Cancel.loc[df_Cancel['CountryName'] == 'Canada', 'Base Rate'] = 'RACK'
    df_Cancel.loc[df_Cancel['CountryName'] == 'Sweden', 'Base Rate'] = 'RACK'
    df_Main.loc[df_Main['CountryName'] != 'United States', 'BWIRateCode'] = bwi
    df_Main.loc[df_Main['CountryName'] == 'United States', 'BWIRateCode'] = bwi
    df_Main.loc[df_Main['AcceptedRateType'] != 'LRA', 'RCFA'] = 'None'
    df_Main.loc[df_Main['AcceptedRateType'] == 'LRA', 'RCFA'] = 'RACK'

    Hotel_Details = pd.concat([df_Main['CRSHotelID'], df_Main['BWIRateCode'], df_Main['RCFA']], axis=1)
    # print(Hotel_Details)
    idx = 0
    Hotel_Details.insert(loc=idx, column='3H', value=3)
    if lra == '0':
        Hotel_Details = Hotel_Details[Hotel_Details.RCFA != 'None']
    Resort_Num = Hotel_Details['CRSHotelID']
    if lra == '1':
        Hotel_Details = Hotel_Details[Hotel_Details.RCFA == 'None']
        Hotel_Details.loc[Hotel_Details['RCFA'] == 'None', 'RCFA'] = bwi
    Resort_Num = Hotel_Details['CRSHotelID']
    # print(Hotel_Details)

    #################################################################################################################################################

    ##############################################################     RATES FOR RESORTS     ###################################################
    sea = []
    lens = []
    res = []
    if rt == '0':
        df_Main = df_Main[df_Main.RoomTypeName == 'STANDARD']
    if rt == '1':
        df_Main = df_Main[df_Main.RoomTypeName != 'STANDARD']
    # print(df_Main)
    for i in range(5):
        sea.append(df_Main[df_Main["Season" + str(i + 1) + "End"].notnull()])

        if sea[i].empty:
            res.append(pd.DataFrame())
        else:
            lens.append(sea[i]['RoomDescription'].str.split(';').map(len))

            res.append(pd.DataFrame({'Resort': np.repeat(sea[i]['CRSHotelID'], lens[i]),
                                     'BWI Rate Code': np.repeat(sea[i]['BWIRateCode'], lens[i]),
                                     'Room Type': chainer(sea[i]['RoomDescription']),
                                     'Begin Date': np.repeat(sea[i]['Season' + str(i + 1) + 'Start'], lens[i]),
                                     'End Date': np.repeat(sea[i]['Season' + str(i + 1) + 'End'], lens[i])
                                     }))

            '''
            #res[i]=res[i].reset_index()
            #res[i]=res[i].drop(columns=['index'])
            print(res[i])
            res[i]['count'] = res[i].groupby('Resort')['Resort'].transform('count')
            count_df=res[i][['Resort','count']]
            count_df=count_df.drop_duplicates()
            print(count_df)
            sea[i]['count']=count_df['count']
            new1 = pd.DataFrame([sea[i].ix[idx] 
                       for idx in sea[i].index 
                       for _ in range(sea[i].ix[idx]['count'])]).reset_index(drop=True)

            #sea[i]=new1
            print(sea[i])
            '''
            if lra == '0':
                if ol == '0':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['2 Person Rate'] = 0
                if ol == '1':
                    res[i]['1 Person Rate'] = 0
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_DBL'], lens[i])
                if ol == '2':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_DBL'], lens[i])

            if lra == '1':
                if ol == '0':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['2 Person Rate'] = 0
                if ol == '1':
                    res[i]['1 Person Rate'] = 0
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_DBL'], lens[i])
                if ol == '2':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_DBL'], lens[i])
            if lra == '2':
                if ol == '0':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['1 Person Rate'].fillna(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_SGL'], inplace=True)
                    res[i]['2 Person Rate'] = 0
                if ol == '1':
                    res[i]['1 Person Rate'] = 0
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_DBL'], lens[i])
                    res[i]['2 Person Rate'].fillna(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_DBL'], inplace=True)
                if ol == '2':
                    res[i]['1 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_SGL'], lens[i])
                    res[i]['1 Person Rate'].fillna(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_SGL'], inplace=True)
                    res[i]['2 Person Rate'] = np.repeat(sea[i]['Preferred_LRA_S' + str(i + 1) + '_DBL'], lens[i])
                    res[i]['2 Person Rate'].fillna(sea[i]['Preferred_Non_LRA_S' + str(i + 1) + '_DBL'], inplace=True)

            res[i]['CountryName'] = np.repeat(sea[i]['CountryName'], lens[i])
            # res[i]['1 Person Rate'].fillna(sea[i]['Preferred_Non_LRA_S'+str(i+1)+'_SGL'], inplace=True)
            try:
                res[i]['Begin Date'] = res[i]['Begin Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            except:
                res[i]['Begin Date'] = pd.to_datetime(res[i]['Begin Date'])
                res[i]['Begin Date'] = res[i]['Begin Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            try:
                res[i]['End Date'] = res[i]['End Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            except:
                res[i]['End Date'] = pd.to_datetime(res[i]['End Date'])
                res[i]['End Date'] = res[i]['End Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
    Rates = pd.concat([res[0], res[1], res[2], res[3], res[4]])
    Rates['colFromIndex'] = Rates.index
    Rates = Rates.sort_values(by=['colFromIndex', 'End Date'])
    Rates = Rates.drop(['colFromIndex'], axis=1)
    RI = pd.DataFrame(Rates)
    idx = 0
    Rates = Rates.drop(columns=['CountryName'])
    Rates.insert(loc=idx, column='5H', value=5)
    # Rates = Rates.dropna(axis=0, subset=['1 Person Rate'])
    Rates = Rates[Rates['Resort'].isin(Resort_Num)]

    ##########################################################################################################################################################
    try:
        df_Blackout = df_Blackout[df_Blackout["BlackoutDateFrom_1"].notnull()]
    except:
        df_Blackout = pd.DataFrame()
        pass

    if rt == '2':
        ids = df_Main["CRSHotelID"]
        dup_res = df_Main[ids.isin(ids[ids.duplicated()])]
        drs = dup_res["CRSHotelID"].unique().tolist()
        for i in drs:
            if df_FairDates.empty:
                pass
            else:
                df_fair_dup = df_FairDates[df_FairDates['PROPCODE'] == i]
                df_FairDates = df_FairDates.append(df_fair_dup, sort=False)
            if df_Blackout.empty:
                pass
            else:
                df_black_dup = df_Blackout[df_Blackout['CRSHotelID'] == i]
                df_Blackout = df_Blackout.append(df_black_dup, sort=False)
        df_FairDates = df_FairDates.reset_index()
        df_Blackout = df_Blackout.reset_index()

    # create duplicates in blackout df
    # create duplicates in Fairdates df

    ##############################################################      FAIR DATE RATES      ############################################################
    seafd = []
    dfFMS = []
    lensfs = []
    refd = []
    refde = pd.DataFrame()

    # print(df_FairDates)
    for i in range(10):
        seafd.append(df_FairDates[df_FairDates["BD" + str(i + 1) + "_END"].notnull()])
        seafd[i] = seafd[i].rename(columns={'PROPCODE': 'CRSHotelID'})
        dfFMS.append(df_Main.merge(seafd[i], on=['CRSHotelID']))
        dfFMS[i] = dfFMS[i].drop_duplicates()
        # print(seafd[i])
        if seafd[i].empty:
            u = 0
            refd.append(pd.DataFrame())
        elif dfFMS[i].empty:
            u = 1
            refd.append(pd.DataFrame())
            string = "Resort Number in Fair Dates Tab does not Match with Main Tab"
            data = pd.DataFrame({"Error": string}, index=[0])
            valid = valid.append(data)
            e = 2
            return (error(valid, e))
        else:
            u = 0
            dfToList = dfFMS[i]['CRSHotelID'].tolist()
            seafd[i] = seafd[i][seafd[i]['CRSHotelID'].isin(dfToList)]
            '''
            if rt=='1':
                seafd[i]=
            '''
            col_ind = seafd[i].index
            dfFMS[i] = dfFMS[i].set_index([col_ind])
            if (dfFMS[i]['RoomDescription'] == 'Upgrade').any():
                lensfs.append(dfFMS[i]['RoomDescription'].str.split(';').map(len))
            # lensfs.append(dfFMS[i]['RoomDescription'].str.split(';').map(len)+len(dfFMS[i][dfFMS[i]['RoomDescription'] == 'Upgrade']))
            else:
                lensfs.append(dfFMS[i]['RoomDescription'].str.split(';').map(len))

            # print(lensfs[i])
            # print(dfFMS[i]['RoomDescription'])
            # print(seafd[i])
            refd.append(pd.DataFrame({'Resort': np.repeat(dfFMS[i]['CRSHotelID'], lensfs[i]),
                                      'BWI Rate Code': np.repeat(dfFMS[i]['BWIRateCode'], lensfs[i]),
                                      'Room Type': chainer(dfFMS[i]['RoomDescription']),
                                      'Begin Date': np.repeat(seafd[i]['BD' + str(i + 1) + '_START'], lensfs[i]),
                                      'End Date': np.repeat(seafd[i]['BD' + str(i + 1) + '_END'], lensfs[i])

                                      }))

            refd[i]['CountryName'] = np.repeat(dfFMS[i]['CountryName'], lensfs[i])
            try:
                refd[i]['Begin Date'] = refd[i]['Begin Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            except:
                refd[i]['Begin Date'] = pd.to_datetime(refd[i]['Begin Date'])
                refd[i]['Begin Date'] = refd[i]['Begin Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            try:
                refd[i]['End Date'] = refd[i]['End Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
            except:
                refd[i]['End Date'] = pd.to_datetime(refd[i]['End Date'])
                refd[i]['End Date'] = refd[i]['End Date'].apply(lambda x: x.strftime('%m/%d/%Y'))

            if ol == '0':
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '1 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT1_SGL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '1 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT2_SGL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '2 Person Rate'] = 0
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '2 Person Rate'] = 0
            if ol == '1':
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '1 Person Rate'] = 0
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '1 Person Rate'] = 0
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '2 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT1_DBL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '2 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT2_DBL'], lensfs[i])
            if ol == '2':
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '1 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT1_SGL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '1 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT2_SGL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] == 'STANDARD', '2 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT1_DBL'], lensfs[i])
                refd[i].loc[dfFMS[i]['RoomTypeName'] != 'STANDARD', '2 Person Rate'] = np.repeat(
                    seafd[i]['BD' + str(i + 1) + '_RT2_DBL'], lensfs[i])
            refd[i]['1 Person Rate'].fillna(0, inplace=True)
            refd[i]['2 Person Rate'].fillna(0, inplace=True)

        refde = pd.concat([refd[i], refde])
    # print(refde)
    ################################################################################################ Blackout Dates without Fair Dates   #############################################################################################
    seabd = []
    dfbMS = []
    lensbs = []
    rebd = []
    bd_id = []
    bd_fd = []
    bd_bd = []
    dupli = pd.DataFrame()
    newd = pd.DataFrame()
    df_FairDates = df_FairDates.rename(columns={'PROPCODE': 'CRSHotelID'})
    print(len(df_Blackout.columns))
    rr = (len(df_Blackout.columns) - 9) / 3
    print(rr)
    rebde = pd.DataFrame()
    for i in range(int(rr)):
        if df_Blackout.empty:
            rebde = pd.DataFrame()
            pass
        else:
            seabd.append(df_Blackout[df_Blackout["BlackoutDateFrom_" + str(i + 1)].notnull()])
            seabd[i] = seabd[i].rename(columns={'PROPCODE': 'CRSHotelID'})
            bd_id.append(pd.DataFrame({'CRSHotelID': seabd[i]['CRSHotelID'][
                ~seabd[i]['CRSHotelID'].isin(df_FairDates['CRSHotelID'])].dropna()}))
            # print(bd_id[i])
            bd_fd.append(df_Main.merge(bd_id[i], on=['CRSHotelID']))
            bd_fd[i] = bd_fd[i].drop_duplicates()
            bd_bd.append(seabd[i].merge(bd_id[i], on=['CRSHotelID']))

            bd_bd[i] = bd_bd[i].merge(bd_fd[i], on=['CRSHotelID'])
            bd_fd[i] = bd_fd[i].reset_index()

            bd_bd[i] = bd_bd[i].drop_duplicates()
            # print(i+1)

            if seabd[i].empty:
                rebd.append(pd.DataFrame())
            elif bd_fd[i].empty:
                rebd.append(pd.DataFrame())

            else:
                if (len(bd_fd[i]) != len(bd_bd[i])):
                    dupli['CRSHotelID'] = bd_fd[i]['CRSHotelID'][bd_fd[i]['CRSHotelID'].duplicated(keep='first')]
                    newd = bd_bd[i].merge(dupli, on=['CRSHotelID'])
                    bd_bd[i] = bd_bd[i].append(newd, ignore_index=True)

                lensbs.append(bd_fd[i]['RoomDescription'].str.split(';').map(len))
                # print(lensbs[i])
                rebd.append(pd.DataFrame({'Resort': np.repeat(bd_fd[i]['CRSHotelID'], lensbs[i]),
                                          'BWI Rate Code': np.repeat(bd_fd[i]['BWIRateCode'], lensbs[i]),
                                          'Room Type': chainer(bd_fd[i]['RoomDescription']),
                                          'Begin Date': np.repeat(bd_bd[i]['BlackoutDateFrom_' + str(i + 1)],
                                                                  lensbs[i]),
                                          'End Date': np.repeat(bd_bd[i]['BlackoutDateTo_' + str(i + 1)], lensbs[i])
                                          }))
                rebd[i]['1 Person Rate'] = 0
                rebd[i]['CountryName'] = np.repeat(bd_fd[i]['CountryName'], lensbs[i])
                rebd[i]['Begin Date'] = rebd[i]['Begin Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
                rebd[i]['End Date'] = rebd[i]['End Date'].apply(lambda x: x.strftime('%m/%d/%Y'))
                rebd[i]['2 Person Rate'] = 0
            rebde = pd.concat([rebd[i], rebde])
    # print(rebde)
    ################################################################################################  Fair Date Rates   #############################################################################################
    if refde.empty:
        fd_Rates = pd.concat([refde, rebde])
        Fair_Rates = pd.DataFrame(fd_Rates)
        if 'CountryName' in Fair_Rates.columns:
            Fair_Rates = fd_Rates.drop(columns=['CountryName'])
        Fair_Rates.insert(loc=idx, column='7H', value=7)
    else:
        fd_Rates = pd.DataFrame(refde)
        fd_Rates['colFromIndex'] = fd_Rates.index
        fd_Rates = fd_Rates.sort_values(by=['colFromIndex', 'End Date'])
        fd_Rates = fd_Rates.drop(['colFromIndex'], axis=1)
        fd_Rates = pd.concat([fd_Rates, rebde])
        idx = 0
        Fair_Rates = fd_Rates.drop(columns=['CountryName'])
        Fair_Rates.insert(loc=idx, column='7H', value=7)

    # print(fd_Rates)
    ################################################################################################ Rate Integrity BD    #####################################################################################################
    C_df = pd.read_excel(path + "//Country_RI.xlsx", sheet_name='Sheet1')
    li = C_df["Country"].tolist()
    if fd_Rates.empty:
        rebdc = pd.DataFrame()
        Fair_Rates = pd.DataFrame()
    else:
        rebdc = fd_Rates.drop(columns=['1 Person Rate'])
        rebdc.loc[rebdc['CountryName'] != 'United States', '1 Person Rate'] = int(ria)
        rebdc.loc[rebdc['CountryName'] == 'United States', '1 Person Rate'] = int(ria)
        for i in li:
            rebdc.loc[rebdc['CountryName'] == i, '1 Person Rate'] = 0
        rebdc.loc[rebdc['CountryName'] != 'United States', '2 Person Rate'] = int(ria)
        rebdc.loc[rebdc['CountryName'] == 'United States', '2 Person Rate'] = int(ria)
        for i in li:
            rebdc.loc[rebdc['CountryName'] == i, '2 Person Rate'] = 0

        rebdc = rebdc.drop(columns=['CountryName'])
        idx = 0
        rebdc.insert(loc=idx, column='6H', value=6)
        new_order = [0, 1, 2, 3, 4, 5, 7, 6]
        # new_order = ['4H','Resort','BWI Rate Code','Room Type','Begin Date','End Date','1 Person Rate','2 Person Rate']
        rebdc = rebdc[rebdc.columns[new_order]]
    # print(RI)

    # print(rebdc)

    # print(rebdc)

    ############################################################################################### Rate Integrity ##############################################################################################################

    RI = RI.drop(columns=['1 Person Rate'])
    RI.loc[RI['CountryName'] != 'United States', '1 Person Rate'] = int(ria)

    RI.loc[RI['CountryName'] == 'United States', '1 Person Rate'] = int(ria)
    for i in li:
        RI.loc[RI['CountryName'] == i, '1 Person Rate'] = 0

    RI.loc[RI['CountryName'] != 'United States', '2 Person Rate'] = int(ria)

    RI.loc[RI['CountryName'] == 'United States', '2 Person Rate'] = int(ria)
    for i in li:
        RI.loc[RI['CountryName'] == i, '2 Person Rate'] = 0

    RI = RI.drop(columns=['CountryName'])
    idx = 0
    RI.insert(loc=idx, column='4H', value=4)
    new_order = [0, 1, 2, 3, 4, 5, 7, 6]
    # new_order = ['4H','Resort','BWI Rate Code','Room Type','Begin Date','End Date','1 Person Rate','2 Person Rate']
    RI = RI[RI.columns[new_order]]
    # print(RI)
    ############################################################################################## Cancellation Policy #############################################################################################
    Cancel_Policy = pd.DataFrame(
        {"Resort": df_Cancel['CRSHotelID'], "BWI Rate Code": df_Main['BWIRateCode'], 'Begin Date': today,
         'End Date': '12/31/2099', "Guarantee Policy": 'GTD', "Cancel Code": df_Cancel['Cancellation'],
         "Base Rate": df_Cancel['Base Rate']})
    idx = 0
    Cancel_Policy.insert(loc=idx, column='8H', value=8)
    # Cancel_Policy=Cancel_Policy.dropna()

    ##############################################################################################   Writing to CSV #############################################################################################################

    d = d[d['Resort'].isin(Resort_Num)]
    Hotel_Details = Hotel_Details.drop_duplicates(subset=['CRSHotelID'], keep='first')
    # d=d.drop_duplicates(subset=['Resort'], keep='first')
    RI = RI[RI['Resort'].isin(Resort_Num)]
    if rebdc.empty:
        pass
    else:
        rebdc = rebdc[rebdc['Resort'].isin(Resort_Num)]
    if Fair_Rates.empty:
        pass
    else:
        Fair_Rates = Fair_Rates[Fair_Rates['Resort'].isin(Resort_Num)]
    Cancel_Policy = Cancel_Policy[Cancel_Policy['Resort'].isin(Resort_Num)]
    Cancel_Policy['Cancel Code'] = Cancel_Policy['Cancel Code'].astype(str)
    Cancel_Policy['Resort'] = Cancel_Policy['Resort'].astype(int)
    Cancel_Policy['Resort'] = Cancel_Policy['Resort'].astype(str)
    Cancel_Policy['Hold Code'] = '4PM'
    for index, row in df_vc.iterrows():
        Cancel_Policy.loc[Cancel_Policy['Cancel Code'] == row['CancelValue'], 'Cancel Code'] = row['Map To CPM Value']
    Cancel_Policy.loc[Cancel_Policy['Cancel Code'] == "6PM", 'Hold Code'] = '6PM'
    Cancel_Policy["BWI Rate Code"] = bwi

    Hotel_Details['Base Rate'] = Cancel_Policy['Base Rate'].values
    Cancel_Policy = Cancel_Policy.drop(columns=['Base Rate'])

    ############################################################################################## Trailing Zeroes ##################################################################################################################

    d_Final = pd.DataFrame(d)
    Hotel_Details_Final = pd.DataFrame(Hotel_Details)
    RI_Final = pd.DataFrame(RI)
    Rates_Final = pd.DataFrame(Rates)
    rebdc_Final = pd.DataFrame(rebdc)
    Fair_Rates_Final = pd.DataFrame(Fair_Rates)
    Cancel_Policy_Final = pd.DataFrame(Cancel_Policy)

    ############################YEAR_VALIDATION###################
    qa = []
    repeat = 0
    Rates['Resort'] = Rates['Resort'].astype(str)
    Rates['Room Type'] = Rates['Room Type'].str.strip()
    Rates['CHECK'] = Rates[['Resort', 'Room Type']].apply(lambda x: ','.join(x), axis=1)
    date_valid = Rates[Rates['CHECK'].duplicated(keep=False)]
    date_valid['Begin_Date'] = pd.to_datetime(date_valid['Begin Date'])
    date_valid['End_Date'] = pd.to_datetime(date_valid['End Date'])
    date_valid = date_valid.sort_values(by=['Resort', 'Begin_Date'])
    # valid = pd.DataFrame()
    date_group = date_valid.groupby('CHECK')
    for name, group in date_group:
        group['overlap2'] = (group['End_Date'].shift() - group['Begin_Date'])
        # group['overlap2']=group['overlap2'].fillna(-1)
        group['overlap2'].replace({pd.NaT: timedelta(-1)}, inplace=True)
        group['overlap'] = np.where(group['overlap2'] != timedelta(-1), True, False)
        print(group['overlap'])
        if (group['overlap'] == True).any():
            qa.append(group.Resort.unique())
            repeat = 1
            w = 1
            string = "Resort Number" + group.Resort.unique() + " " + group[
                'Room Type'].unique() + "Has overlapping dates or date gaps in Main Tab"
            data = pd.DataFrame({"Error": string})
            valid = valid.append(data)
            e = 2
    ##################################################
    ##################################################
    ################ Fair Date Validation#############
    if refde.empty:
        pass
    else:
        refde['Resort'] = refde['Resort'].astype(str)
        refde['CHECK'] = refde[['Resort', 'Room Type']].apply(lambda x: ','.join(x), axis=1)
        date_valid = refde[refde['CHECK'].duplicated(keep=False)]
        date_valid['Begin_Date'] = pd.to_datetime(date_valid['Begin Date'])
        date_valid['End_Date'] = pd.to_datetime(date_valid['End Date'])
        date_valid = date_valid.sort_values(by=['Resort', 'Begin_Date'])
        # print(date_valid['CHECK'])
        date_group = date_valid.groupby('CHECK')
        for name, group in date_group:
            group['overlap2'] = (group['End_Date'].shift() - group['Begin_Date'])
            # group['overlap2']=group['overlap2'].fillna(-1)
            group['overlap2'].replace({pd.NaT: timedelta(-1)}, inplace=True)
            group['overlap'] = np.where(group['overlap2'] > timedelta(-1), True, False)
            if (group['overlap'] == True).any():
                qa.append(group.Resort.unique())
                repeat = 1
                w = 1
                string = "Resort Number" + group.Resort.unique() + " " + group[
                    'Room Type'].unique() + "Has overlapping dates in FAIRDATES Tab"
                data = pd.DataFrame({"Error": string}, index=[0])
                valid = valid.append(data)
                e = 2
    ##########################################################################################################################################################################################################
    ####################################################################### Fair Date Validation ##############################################################################################################
    if rebde.empty:
        pass
    else:
        rebde['Resort'] = rebde['Resort'].astype(str)
        rebde['CHECK'] = rebde[['Resort', 'Room Type']].apply(lambda x: ','.join(x), axis=1)
        date_valid = rebde[rebde['CHECK'].duplicated(keep=False)]
        date_valid['Begin_Date'] = pd.to_datetime(date_valid['Begin Date'])
        date_valid['End_Date'] = pd.to_datetime(date_valid['End Date'])
        date_valid = date_valid.sort_values(by=['Resort', 'Begin_Date'])
        # print(date_valid['CHECK'])
        date_group = date_valid.groupby('CHECK')
        for name, group in date_group:
            group['overlap2'] = (group['End_Date'].shift() - group['Begin_Date'])
            # group['overlap2']=group['overlap2'].fillna(-1)
            group['overlap2'].replace({pd.NaT: timedelta(-1)}, inplace=True)
            group['overlap'] = np.where(group['overlap2'] > timedelta(-1), True, False)
            if (group['overlap'] == True).any():
                w = 1
                qa.append(group.Resort.unique())
                repeat = 1
                string = "Resort Number" + group.Resort.unique() + " " + group[
                    'Room Type'].unique() + "Has overlapping dates in BlackoutDates Tab"
                data = pd.DataFrame({"Error": string}, index=[0])
                valid = valid.append(data)
                e = 2
    ##############################################################################################################################################################################################################
    ##################################################### Resort Number and RoomType Validation##################################################################################################################
    ww = []
    df_resort = pd.read_csv(path + '//Resort.csv')
    df_roomtype = pd.read_csv(path + '//Roomtype.csv')
    df_roomtype['RESORT'] = df_roomtype['RESORT'].astype(str)
    # df_roomtype['RESORT'] =df_roomtype['RESORT'].apply(lambda x:x.zfill(5))
    df_roomtype['CHECK'] = df_roomtype[['RESORT', 'NEXUS_ROOM_TYPE']].apply(lambda x: ','.join(x), axis=1)
    Rates = Rates[Rates['Room Type'] != 'Standard']
    Rates = Rates[Rates['Room Type'] != 'Upgrade']
    Rates = Rates[Rates['Room Type'] != 'STANDARD']
    Rates = Rates[Rates['Room Type'] != 'UPGRADE']
    Rates = Rates[Rates['Room Type'] != 'EXISTING']
    Rates = Rates[Rates['Room Type'] != 'ROH']
    Rates = Rates[Rates['Room Type'] != 'ALL']

    Rates['Room Type'] = Rates['Room Type'].str.replace(r',\d+', '')
    Rates['CHECK'] = Rates[['Resort', 'Room Type']].apply(lambda x: ','.join(x), axis=1)
    # Rates['CHECK'] = Rates['CHECK'].str.replace(r',vigo', '')
    Rates['va'] = Rates.CHECK.isin(df_roomtype.CHECK)
    d['va'] = d.Resort.isin(df_resort.RESORT)

    if (d['va'] == False).any():
        string = "The Following Resort Numbers Does not Match"
        dferror = pd.DataFrame({"Error": (d[~d.va].Resort).astype(str)})
        dferror1 = pd.DataFrame({"Error": (d[~d.va].Resort)})
        repeat = 1
        w = 1
        ww.extend(dferror1["Error"].tolist())
        data = pd.DataFrame({"Error": string}, index=[0])
        data = pd.concat([data, dferror], axis=0)
        # data=data.append({"Error": d['Resort'].where(d['va'] == False)},ignore_index=True)
        valid = valid.append(data)
        e = 2

    if (Rates['va'] == False).any():
        string = "The Following Resort Room Type Does not Match"
        dferror = pd.DataFrame({"Error": Rates[~Rates.va].CHECK})
        dferror1 = pd.DataFrame({"Error": (Rates[~Rates.va].Resort)})
        repeat = 1
        w = 1
        ww.extend(dferror1["Error"].tolist())
        data = pd.DataFrame({"Error": string}, index=[0])
        data = pd.concat([data, dferror], axis=0)
        # data=data.append({"Error": d['Resort'].where(d['va'] == False)},ignore_index=True)
        valid = valid.append(data)
        e = 2

    if e == 2:
        error = valid.to_csv(path + '//df.csv', index=False)
    if w == 0:
        d_Final['Resort'] = d_Final['Resort'].astype(int)
        d_Final['Resort'] = d_Final['Resort'].astype(str)
        d_Final['Resort'] = d_Final['Resort'].apply(lambda x: x.zfill(5))
        d_Final = d_Final.drop(columns=['va'])

        Hotel_Details_Final['CRSHotelID'] = Hotel_Details_Final['CRSHotelID'].astype(str)
        Hotel_Details_Final['CRSHotelID'] = Hotel_Details_Final['CRSHotelID'].apply(lambda x: x.zfill(5))

        RI_Final['Resort'] = RI_Final['Resort'].astype(str)
        RI_Final['Resort'] = RI_Final['Resort'].apply(lambda x: x.zfill(5))

        Rates_Final['Resort'] = Rates_Final['Resort'].astype(str)
        Rates_Final['Resort'] = Rates_Final['Resort'].apply(lambda x: x.zfill(5))

        if rebdc_Final.empty:
            rebdc_Final = pd.DataFrame(
                columns=['6H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person Rate',
                         '2 Person Rate'])
            pass
        else:
            rebdc_Final['Resort'] = rebdc_Final['Resort'].astype(str)
            rebdc_Final['Resort'] = rebdc_Final['Resort'].apply(lambda x: x.zfill(5))

        if Fair_Rates_Final.empty:
            Fair_Rates_Final = pd.DataFrame(
                columns=['7H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person Rate',
                         '2 Person Rate'])
            pass
        else:
            Fair_Rates_Final['Resort'] = Fair_Rates_Final['Resort'].astype(str)
            Fair_Rates_Final['Resort'] = Fair_Rates_Final['Resort'].apply(lambda x: x.zfill(5))

        Cancel_Policy_Final['Resort'] = Cancel_Policy_Final['Resort'].astype(str)
        Cancel_Policy_Final['Resort'] = Cancel_Policy_Final['Resort'].apply(lambda x: x.zfill(5))
        ##################################################### Adding Missing Columns and Updating Column names ###########################################################################

        Hotel_Details_Final['Flat/Percentage'] = 'P'
        Hotel_Details_Final.columns = ['3H', 'Resort', 'BWI Rate Codes', 'Rate Code for Avail', 'Base Rate',
                                       'Flat/Percentage']
        RI_Final.columns = ['4H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person rate',
                            '2 Person rate']
        Rates_Final.columns = ['5H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person rate',
                               '2 Person rate']
        rebdc_Final.columns = ['6H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person rate',
                               '2 Person rate']
        rebdc_Final = rebdc_Final.sort_values(by=['Resort', 'End Date'])
        Fair_Rates_Final.columns = ['7H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date',
                                    '1 Person rate', '2 Person rate']
        Fair_Rates_Final = Fair_Rates_Final.sort_values(by=['Resort', 'End Date'])
        H1['Corp Acct#'] = H1['Corp Acct#'].astype(str)
        H1['Corp Acct#'] = H1['Corp Acct#'].apply(lambda x: x.zfill(8))
        d_Final['Corp Acct #'] = d_Final['Corp Acct #'].astype(str)
        d_Final['Corp Acct #'] = d_Final['Corp Acct #'].apply(lambda x: x.zfill(8))
        list_dow = ['S', 'M', 'T', 'W', 'T', 'F', 'S']
        RI_Final['1 Person rate'] = RI_Final['1 Person rate'] / (-100)
        RI_Final['2 Person rate'] = RI_Final['2 Person rate'] / (-100)

        rebdc_Final['1 Person rate'] = rebdc_Final['1 Person rate'] / (-100)
        rebdc_Final['2 Person rate'] = rebdc_Final['2 Person rate'] / (-100)

        for j in list_dow:
            RI_Final[j] = None
            Rates_Final[j] = None
            rebdc_Final[j] = None
            Fair_Rates_Final[j] = None

        ##make changes here
        d_Final = d_Final.drop(columns=['Begin Date', 'End Date'])
        d_Final['Inactive Date'] = None

        newdict = {'Sabre': 'AA', 'WorldSpan': '1P', 'Amadeus': '1A', 'Galileo': 'UA', 'Web': 'WB'}
        list_new = ['Sabre', 'WorldSpan', 'Amadeus', 'Galileo', 'Web']
        new_df_gds = pd.DataFrame()
        d_Final.reset_index(inplace=True)

        for i, row in d_Final.iterrows():
            gds_y_list = []
            for gds in list_new:
                if row[gds] == 'Y':
                    gds_y_list.append(gds)
            df_gds = pd.concat([d_Final.iloc[i]] * len(gds_y_list), axis=1, ignore_index=True)
            df_gds = df_gds.T
            df_gds['GDS HOST'] = gds_y_list
            df_gds = df_gds.replace({"GDS HOST": newdict})
            new_df_gds = pd.concat([new_df_gds, df_gds])

        d_Final = pd.DataFrame(new_df_gds)
        d_Final.reset_index(inplace=True)
        Hotel_Details_Final['RATE LEVEL'] = 'B'
        Hotel_Details_Final.loc[Hotel_Details_Final['Rate Code for Avail'] == 'RACK', 'RATE LEVEL'] = 'A'
        Hotel_Details_Final['FORECAST SEGMENT'] = 'CORPORATE'
        Hotel_Details_Final['Rate Code for Confirm'] = None
        dup_d_final = d_Final.drop_duplicates('Resort')
        Hotel_Details_Final.set_index(['Resort'], inplace=True)
        dup_d_final.set_index(['Resort'], inplace=True)
        Hotel_Details_Final['ROH_YN'] = dup_d_final['Map to ROH (Y/N)']
        Hotel_Details_Final.reset_index(inplace=True)
        Hotel_Details_Final['MODIFY_RATE_YN'] = 'N'
        Hotel_Details_Final['MODIFY_CONFIRM_YN'] = 'N'
        Hotel_Details_Final['MODIFY_INVENTORY_YN'] = 'N'
        Hotel_Details_Final['MODIFY_STATUS_YN'] = 'N'

        d_Final = d_Final.drop(columns=list_new)
        d_Final = d_Final.drop(columns=['Map to ROH (Y/N)', 'index'])

        print(Hotel_Details_Final)
        Hotel_Details_Final = Hotel_Details_Final[
            ['3H', 'Resort', 'BWI Rate Codes', 'RATE LEVEL', 'Rate Code for Avail', 'Base Rate', 'FORECAST SEGMENT',
             'Flat/Percentage', 'Rate Code for Confirm', 'MODIFY_RATE_YN', 'MODIFY_CONFIRM_YN', 'MODIFY_INVENTORY_YN',
             'MODIFY_STATUS_YN', 'ROH_YN']]
        d_Final = d_Final[
            ['2H', 'Corp Acct #', 'GDS HOST', 'GDS Rate Codes', 'Multi Rate Codes', 'Resort', 'BWI Rate Code',
             'Inactive Date']]

        ####################################################################################################################################################
        writer = pd.ExcelWriter(path + '//Pd.xlsx', engine='xlsxwriter')

        # Position the dataframes in the worksheet.
        H1.to_excel(writer, sheet_name='Sheet1', index=False)  # Default position, cell A1.
        d_Final.to_excel(writer, sheet_name='Sheet1', startrow=len(H1) + 1, index=False)
        Hotel_Details_Final.to_excel(writer, sheet_name='Sheet1', startrow=len(H1) + len(d_Final) + 2, index=False)
        RI_Final.to_excel(writer, sheet_name='Sheet1', startrow=len(H1) + len(d_Final) + len(Hotel_Details_Final) + 3,
                          index=False)  # Default position, cell A1.
        Rates_Final.to_excel(writer, sheet_name='Sheet1',
                             startrow=len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + 4,
                             index=False)
        rebdc_Final.to_excel(writer, sheet_name='Sheet1',
                             startrow=len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + len(
                                 Rates_Final) + 5, index=False)
        Fair_Rates_Final.to_excel(writer, sheet_name='Sheet1',
                                  startrow=len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + len(
                                      Rates_Final) + len(rebdc_Final) + 6, index=False)
        Cancel_Policy_Final.to_excel(writer, sheet_name='Sheet1',
                                     startrow=len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + len(
                                         Rates_Final) + len(rebdc_Final) + len(Fair_Rates_Final) + 7, index=False)
        df_h1.to_excel(writer, sheet_name='Section 2H', index=False)
        df_h2.to_excel(writer, sheet_name='Section 2H', index=False, startrow=len(df_h1) + 1)
        writer.save()

        wb = openpyxl.load_workbook(path + '//Pd.xlsx')
        ws = wb.active
        row_num = []
        for cell in ws['L']:
            if (cell.value is not None):  # We need to check that the cell is not empty.
                if 'W' in cell.value:  # Check if the value of the cell contains the text 'W'
                    # print(cell.row)
                    ws['M' + str(cell.row)].value = 'T'
                    ws['N' + str(cell.row)].value = 'F'
                    ws['O' + str(cell.row)].value = 'S'

        for i in range(len(H1) + 3, len(H1) + len(d_Final) + 3):
            if ws['F' + str(i)].value is None or ws['F' + str(i)].value == 'N':
                ws['F' + str(i)].fill = openpyxl.styles.PatternFill('solid', openpyxl.styles.colors.GREEN)
        for i in range(len(H1) + 3,
                       len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + len(Rates_Final) + len(
                               rebdc_Final) + len(Fair_Rates_Final) + len(Cancel_Policy_Final) + 9):
            if ws['C' + str(i)].value is None or ws['C' + str(i)].value == 'N':
                ws['C' + str(i)].fill = openpyxl.styles.PatternFill('solid', openpyxl.styles.colors.GREEN)
        for i in range(len(H1) + 3,
                       len(H1) + len(d_Final) + len(Hotel_Details_Final) + len(RI_Final) + len(Rates_Final) + len(
                               rebdc_Final) + len(Fair_Rates_Final) + len(Cancel_Policy_Final) + 9):
            if ws['D' + str(i)].value is None or ws['D' + str(i)].value == 'N':
                ws['D' + str(i)].fill = openpyxl.styles.PatternFill('solid', openpyxl.styles.colors.GREEN)

        wb.save(path + '//Pd.xlsx')

        # print(Cancel_Policy_Final.iloc[2]['Cancel Code'])

        #################################################################################################################################################################################################################################

        user = d_Final.to_csv(None, index=False).encode()
        CRS = H1.to_csv(None, index=False).encode()
        Hotel = Hotel_Details_Final.to_csv(None, index=False).encode()
        Rate_Integrity = RI_Final.to_csv(None, index=False).encode()
        Room_rate = Rates_Final.to_csv(None, index=False).encode()
        if rebdc_Final.empty:
            rebdc_Final = pd.DataFrame(
                columns=['6H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person Rate',
                         '2 Person Rate'])
            pass
        else:
            RIBD = rebdc_Final.to_csv(None, index=False).encode()
        if Fair_Rates_Final.empty:
            Fair_Rates_Final = pd.DataFrame(
                columns=['7H', 'Resort', 'BWI Rate Code', 'Room Type', 'Begin Date', 'End Date', '1 Person Rate',
                         '2 Person Rate'])
            pass
        else:
            Rate_Fair = Fair_Rates_Final.to_csv(None, index=False).encode()
        Hotel_cancel = Cancel_Policy_Final.to_csv(None, index=False).encode()
        # print(user.user())
        # fs = s3fs.S3FileSystem(key="AKIAIIX672H64FRKESLA", secret="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")
        with open(path + '//file.csv', 'wb') as f:
            f.write(CRS)
            f.write(user)
            f.write(Hotel)

            f.write(Rate_Integrity)
            f.write(Room_rate)
            if Fair_Rates.empty:
                pass
            else:
                f.write(RIBD)
            if Fair_Rates.empty:
                pass
            else:
                f.write(Rate_Fair)
            f.write(Hotel_cancel)
        ######################################################Excel Writer##################################################################################################

    ###########################################################################################################################################
    ################################################ ERROR ####################################################################################
    ############################################################################################################################################
    if repeat != 0:
        return (etl(fd, rt, ol, lra, ria, bwi, repeat, qa, ww))

    return e
###########################################################################################################################################
################################################ ERROR ####################################################################################


############################################################################################################################################