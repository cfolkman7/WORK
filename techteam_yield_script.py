# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 14:14:59 2020

@author: connor.folkman
"""

import pandas as pd
from pandas import ExcelWriter, DataFrame

month = ' '
month2 = ' '

REV_B_PN_WOS = [502438001, 502438002, 502438003]

REV_D_PN_WOS = [504365001, 504365002, 504365003]

REV_B_PN_FA = [502724001, 502724002, 502724003, 502724004, 502724005, 502724006, 502724007, 502724008, 502724009,
               502725001, 502725002, 502725003, 502725004, 502725005, 502725006, 502725007, 502725008, 502725009,
               502727001, 502727002, 502727003, 502727004, 502727005, 502727006, 502727007, 502727008, 502727009]

REV_D_PN_FA = [504367001, 504367002, 504367003, 504367004, 504367005, 504367006, 504367007, 504367008, 504367009,
               504368001, 504368002, 504368003, 504368004, 504368005, 504368006, 504368007, 504368008, 504368009,
               504369001, 504369002, 504369003, 504369004, 504369005, 504369006, 504369007, 504369008, 504369009]

PURSUE_PN = [504281001, 504281002, 504281003, 504281004, 504281005, 504281006, 504281007, 504281008, 504281009,
             504280001, 504280002, 504280003, 504280004, 504280005, 504280006, 504280007, 504280008, 504280009]

FA_21F_PN = [504230001, 504230002, 504230003, 504230004, 504230005, 504230006, 504230007, 504230008, 504230009]

def parse_yield(part_num, catalog_num, yield_num_org, yield_num_clo, month_abr, month):
    i=0
    k=0
    sum_org = 0
    sum_clo = 0
    yield_list_org = []
    yield_list_clo = []
    total_yield = 0.0
    while i < len(catalog_num):
        j=0
        while j < len(part_num):
            if catalog_num[i] == part_num[j]:
                if (yield_num_org[i] > 0) & (yield_num_clo[i] > 0):
                    yield_list_org.append(yield_num_org[i])
                    yield_list_clo.append(yield_num_clo[i])
            j+=1
        i+=1
    while k < len(yield_list_org):
        sum_org += yield_list_org[k]
        k+=1
    k=0
    while k < len(yield_list_clo):
        sum_clo += yield_list_clo[k]
        k+=1
    total_yield = (round((sum_clo/sum_org), 3) * 100)
    return total_yield

def month_finder(month_num):
    i = 4
    while i < len(month_num):
        if (isinstance(month_num[i], str) == False) & ((month_num[i] == 'NaT') == False):
            month_abr = str(month_num[i].strftime('%b'))
            break
        i+=1
    return month_abr
 
data_to_sparse = pd.read_excel('//USSJFS1/Tech_Team_Files/Team # 18 Maestro_Pursue/2020 Presentations/YIELD_DATA_TO_PARSE.xlsx', sheet_name = 'MAESTRO', header=3, mangle_dupe_cols=True)

catalog_wos_list = data_to_sparse['Number'].to_list()
catalog_fa_list = data_to_sparse['Number.1'].to_list()
catalog_wos_list2 = data_to_sparse['Number.2'].to_list()
catalog_fa_list2 = data_to_sparse['Number.3'].to_list()
yield_wos_list_org = data_to_sparse['Original'].to_list()
yield_wos_list_clo = data_to_sparse['Closed'].to_list()
yield_fa_list_org = data_to_sparse['Original.1'].to_list()
yield_fa_list_clo = data_to_sparse['Closed.2'].to_list()
yield_wos_list_org2 = data_to_sparse['Original.2'].to_list()
yield_wos_list_clo2 = data_to_sparse['Closed.4'].to_list()
yield_fa_list_org2 = data_to_sparse['Original.3'].to_list()
yield_fa_list_clo2 = data_to_sparse['Closed.6'].to_list()
month_of_yield = data_to_sparse['Due'].to_list()
month_of_yield2 = data_to_sparse['Due.2'].to_list()

df_parsed_yield_data = DataFrame([[month_finder(month_of_yield), 'Yield %'], 
                                  ['REV B WOS YIELD', parse_yield(REV_B_PN_WOS, catalog_wos_list, yield_wos_list_org, yield_wos_list_clo, month_finder(month_of_yield), month_of_yield)], 
                                  ['REV D WOS YIELD', parse_yield(REV_D_PN_WOS, catalog_wos_list, yield_wos_list_org, yield_wos_list_clo, month_finder(month_of_yield), month_of_yield)],
                                  ['REV B FA YIELD', parse_yield(REV_B_PN_FA, catalog_fa_list, yield_fa_list_org, yield_fa_list_clo, month_finder(month_of_yield), month_of_yield)], 
                                  ['REV D FA YIELD', parse_yield(REV_D_PN_FA, catalog_fa_list, yield_fa_list_org, yield_fa_list_clo, month_finder(month_of_yield), month_of_yield)],
                                  ['PURSUE YIELD', parse_yield(PURSUE_PN, catalog_fa_list, yield_fa_list_org, yield_fa_list_clo, month_finder(month_of_yield), month_of_yield)], 
                                  ['2.1F FA YIELD', parse_yield(FA_21F_PN, catalog_fa_list, yield_fa_list_org, yield_fa_list_clo, month_finder(month_of_yield), month_of_yield)]])

df_parsed_yield_data2 = DataFrame([[month_finder(month_of_yield2), 'Yield %'], 
                                   ['REV B WOS YIELD', parse_yield(REV_B_PN_WOS, catalog_wos_list2, yield_wos_list_org2, yield_wos_list_clo2, month_finder(month_of_yield2), month_of_yield2)], 
                                   ['REV D WOS YIELD', parse_yield(REV_D_PN_WOS, catalog_wos_list2, yield_wos_list_org2, yield_wos_list_clo2, month_finder(month_of_yield2), month_of_yield2)],
                                   ['REV B FA YIELD', parse_yield(REV_B_PN_FA, catalog_fa_list2, yield_fa_list_org2, yield_fa_list_clo2, month_finder(month_of_yield2), month_of_yield2)], 
                                   ['REV D FA YIELD', parse_yield(REV_D_PN_FA, catalog_fa_list2, yield_fa_list_org2, yield_fa_list_clo2, month_finder(month_of_yield2), month_of_yield2)],
                                   ['PURSUE YIELD', parse_yield(PURSUE_PN, catalog_fa_list2, yield_fa_list_org2, yield_fa_list_clo2, month_finder(month_of_yield2), month_of_yield2)], 
                                   ['2.1F FA YIELD', parse_yield(FA_21F_PN, catalog_fa_list2, yield_fa_list_org2, yield_fa_list_clo2, month_finder(month_of_yield2), month_of_yield2)]])

writer = ExcelWriter('//USSJFS1/Tech_Team_Files/Team # 18 Maestro_Pursue/2020 Presentations/PARSED_YIELD_DATA.xlsx')
df_parsed_yield_data.to_excel(writer, month_finder(month_of_yield), index=False)
df_parsed_yield_data2.to_excel(writer, month_finder(month_of_yield2), index=False)
writer.save()
