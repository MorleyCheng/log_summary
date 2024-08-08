# -*- coding: utf-8 -*-
from pandas import DataFrame,ExcelWriter,concat
from openpyxl import load_workbook
#import pandas as pd
import os,shutil
from chardet import detect


def log_summary_v05(file):
    rw_bin1 = ['rw_0:0000A000,0000A000', 'rw_1:0000A000,0000A000', 'rw_2:0000A000,0000A000', 'rw_3:0000A000,0000A000', 'rw_4:0000A000,0000A000', 'rw_5:0000A000,0000A000', 'rw_6:0000A000,0000A000', 'rw_7:0000A000,0000A000', 'rw_8:0000A000,0000A000', 'rw_9:0000A000,0000A000', 'rw_10:0000A000,0000A000', 'rw_11:0000A000,0000A000', 'rw_12:0000A000,0000A000', 'rw_13:0000A000,0000A000', 'rw_14:0000A000,0000A000', 'rw_15:0000A000,0000A000', 'rw_16:0000A000,0000A000', 'rw_17:0000A000,0000A000', 'rw_18:0000A000,0000A000', 'rw_19:0000A000,0000A000', 'rw_20:00000000,001C2000', 'rw_21:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        FTbin = ''
        FTsite = ''
        err = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')
            except:
                pass

            if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                dic.update({line1[0]:err})
            elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                dic.update({line1[0]:''})
    ###### V06
            #if line.startswith('UHS1 scan'):
                #dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])
                if line.startswith('rw_21'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('FtBin'):
                FTbin = line[5:-7].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    #outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    #outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    #outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf



def log_summary_v06(file):
    rw_bin1 = ['rw_0:0000A000,0000A000', 'rw_1:0000A000,0000A000', 'rw_2:0000A000,0000A000', 'rw_3:0000A000,0000A000', 'rw_4:0000A000,0000A000', 'rw_5:0000A000,0000A000', 'rw_6:0000A000,0000A000', 'rw_7:0000A000,0000A000', 'rw_8:0000A000,0000A000', 'rw_9:0000A000,0000A000', 'rw_10:0000A000,0000A000', 'rw_11:0000A000,0000A000', 'rw_12:0000A000,0000A000', 'rw_13:0000A000,0000A000', 'rw_14:0000A000,0000A000', 'rw_15:0000A000,0000A000', 'rw_16:0000A000,0000A000', 'rw_17:0000A000,0000A000', 'rw_18:0000A000,0000A000', 'rw_19:0000A000,0000A000', 'rw_20:00000000,001C2000', 'rw_21:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        FTbin = ''
        FTsite = ''
        err = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')
            except:
                pass


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})
  

    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])  #+','+line1[7]
                if line.startswith('rw_21'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('FtBin'):
                FTbin = line[5:-1].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf
    


def log_summary_7449_v03(file):
    rw_bin1 = ['rw_0:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        end_time = ''
        FTbin = ''
        FTsite = ''
        err = ''
        site1_start_time = ''
        site2_start_time = ''
        site3_start_time = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')

            except:
                pass

            #0206
            if line.startswith('Start timer count') and line1[-1] == 'site:0':
                site1_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('Start timer count') and line1[-1] == 'site:1':
                site2_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('Start timer count') and line1[-1] == 'site:2':
                site3_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            ##


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})

    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])
                if line.startswith('rw_0'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('End FT Test'):  #0206
                end_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()   #0206

            if line.startswith('FtBin'):
                FTbin = line[5:-1].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Start_time'] = site1_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time                    #0206
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Start_time'] = site2_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Start_time'] = site3_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf
    

def log_summary_7494_FT3_v02(file):
    rw_bin1 = ['rw_0:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        FTbin = ''
        FTsite = ''
        err = ''
        site1_start_time = ''
        site2_start_time = ''
        site3_start_time = ''
        end_time = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')

            except:
                pass

            #0206
            if line.startswith('Start timer count') and line1[-1] == 'site:0':
                site1_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('Start timer count') and line1[-1] == 'site:1':
                site2_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('Start timer count') and line1[-1] == 'site:2':
                site3_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            ##


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})

    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])
                if line.startswith('rw_0'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('End FT Test'):  #0206
                end_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()   #0206

            if line.startswith('FtBin'):
                FTbin = line[5:-1].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Start_time'] = site1_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time                    #0206
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Start_time'] = site2_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Start_time'] = site3_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf



def log_summary_7428_v07(file):
    rw_bin1 = ['rw_0:0000A000,0000A000', 'rw_1:0000A000,0000A000', 'rw_2:0000A000,0000A000', 'rw_3:0000A000,0000A000', 'rw_4:0000A000,0000A000', 'rw_5:0000A000,0000A000', 'rw_6:0000A000,0000A000', 'rw_7:0000A000,0000A000', 'rw_8:0000A000,0000A000', 'rw_9:0000A000,0000A000', 'rw_10:0000A000,0000A000', 'rw_11:0000A000,0000A000', 'rw_12:0000A000,0000A000', 'rw_13:0000A000,0000A000', 'rw_14:0000A000,0000A000', 'rw_15:0000A000,0000A000', 'rw_16:0000A000,0000A000', 'rw_17:0000A000,0000A000', 'rw_18:0000A000,0000A000', 'rw_19:0000A000,0000A000', 'rw_20:00000000,001C2000', 'rw_21:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        end_time = ''
        FTbin = ''
        FTsite = ''
        err = ''
        site1_start_time = ''
        site2_start_time = ''
        site3_start_time = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')

            except:
                pass

            #0206
            if line.startswith('[Timer] Start') and line1[-1] == 'site:0':
                site1_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:1':
                site2_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:2':
                site3_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            ##


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})

    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6]+','+line1[7])
                if line.startswith('rw_21'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('[Arduino] EOT'):  #0206
                end_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()   #0206

            if line.startswith('FtBin'):
                FTbin = line[5:-1].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Start_time'] = site1_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time                    #0206
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Start_time'] = site2_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Start_time'] = site3_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf


def log_summary_7494_FT3_v03(file):
    rw_bin1 = ['rw_0:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        FTbin = ''
        FTsite = ''
        err = ''
        site1_start_time = ''
        site2_start_time = ''
        site3_start_time = ''
        end_time = ''
        list_f = list(f)
        #print(list_f)
        count = 0
        
        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')

            except:
                pass

            #0206
            if line.startswith('[Timer] Start') and line1[-1] == 'site:0':
                site1_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:1':
                site2_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:2':
                site3_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            ##


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})

    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])
                if line.startswith('rw_0'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('[Arduino] EOT'):  #0206
                end_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()   #0206

            if line.startswith('FtBin'):
                FTbin = line[5:-1].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Start_time'] = site1_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time                    #0206
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Start_time'] = site2_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Start_time'] = site3_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06
                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

        f.close()
        outdf['File'] = file
        return outdf


def log_summary_7428_FT3_v08_2(file):
    rw_bin1 = ['rw_0:00000800,00000800']
    outdf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time','1_SMbus','1_DUT','1_L1_count','2_SMbus','2_DUT','2_L1_count','LotID','Runcard','LoadBoard']) 
    outdf.index = outdf['SN']
    with open(file,encoding='utf-16-le') as f:
        #enco = chardet.detect(f.read())
        #f.read().decode(enco['encoding'])
        dic = {}
        rw_lst = []
        test_time =''
        FTbin = ''
        FTsite = ''
        err = ''
        site1_start_time = ''
        site2_start_time = ''
        site3_start_time = ''
        site4_start_time = ''
        end_time = ''
        list_f = list(f)
        #print(list_f)
        LotID = ''
        Runcard = ''
        LoadBoard = ''

        count = 0
        ft_on = 0

        site1_1_smbus = ''
        site1_1_dut = ''
        site1_1_l1_count = ''
        site1_2_smbus = ''
        site1_2_dut = ''
        site1_2_l1_count = ''

        site2_1_smbus = ''
        site2_1_dut = ''
        site2_1_l1_count = ''
        site2_2_smbus = ''
        site2_2_dut = ''
        site2_2_l1_count = ''

        site3_1_smbus = ''
        site3_1_dut = ''
        site3_1_l1_count = ''
        site3_2_smbus = ''
        site3_2_dut = ''
        site3_2_l1_count = ''

        site4_1_smbus = ''
        site4_1_dut = ''
        site4_1_l1_count = ''
        site4_2_smbus = ''
        site4_2_dut = ''
        site4_2_l1_count = ''

        for i in range(0,len(list_f)):
            #print(list_f[i])
            try:
                line = list_f[i].strip().split("****")[1].strip().split(':')[1].strip()
                line1 = list_f[i].strip().split("****")[1].strip().split(' ')

            except:
                pass
            
            if line.startswith('** LotID'):
                LotID = line1[2].split(':')[1][:-1]
                Runcard = line1[3].split(':')[1][:-1]
                LoadBoard = line1[4].split(':')[1]                
#### L1            
            if line.startswith('<====== Ft On'):
                ft_on += 1

## SM bus, first pwr cycle
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS0:' and ft_on ==1:
                site1_1_smbus = line1[-1]                
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS1:' and ft_on ==1:
                site2_1_smbus = line1[-1] 
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS2:' and ft_on ==1:
                site3_1_smbus = line1[-1] 
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS3:' and ft_on ==1:
                site4_1_smbus = line1[-1] 


## Set DUT to L1, first pwr cycle
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS0:' and ft_on ==1:
                site1_1_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS1:' and ft_on ==1:
                site2_1_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS2:' and ft_on ==1:
                site3_1_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS3:' and ft_on ==1:
                site4_1_dut = line1[-1] 

## L1 count, first pwr cycle
            if line.startswith('L1 test count') and line1[0] == 'DbgS0:' and ft_on ==1:
                site1_1_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS1:' and ft_on ==1:
                site2_1_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS2:' and ft_on ==1:
                site3_1_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS3:' and ft_on ==1:
                site4_1_l1_count = line1[-1] 


## SM bus, second  pwr cycle
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS0:' and ft_on ==2:
                site1_2_smbus = line1[-1]                
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS1:' and ft_on ==2:
                site2_2_smbus = line1[-1] 
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS2:' and ft_on ==2:
                site3_2_smbus = line1[-1] 
            if line.startswith('SBUS N=1 [05] =>') and line1[0] == 'DbgS3:' and ft_on ==2:
                site4_2_smbus = line1[-1] 


## Set DUT to L1, second pwr cycle
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS0:' and ft_on ==2:
                site1_2_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS1:' and ft_on ==2:
                site2_2_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS2:' and ft_on ==2:
                site3_2_dut = line1[-1] 
            if line.startswith('PCI[090] =>') and line1[0] == 'DbgS3:' and ft_on ==2:
                site4_2_dut = line1[-1] 

## L1 count, second pwr cycle
            if line.startswith('L1 test count') and line1[0] == 'DbgS0:' and ft_on ==2:
                site1_2_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS1:' and ft_on ==2:
                site2_2_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS2:' and ft_on ==2:
                site3_2_l1_count = line1[-1] 
            if line.startswith('L1 test count') and line1[0] == 'DbgS3:' and ft_on ==2:
                site4_2_l1_count = line1[-1] 


####


            #0206
            if line.startswith('[Timer] Start') and line1[-1] == 'site:0':
                site1_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:1':
                site2_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:2':
                site3_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            if line.startswith('[Timer] Start') and line1[-1] == 'site:3':
                site4_start_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()
            ##


            if len(line1) > 3:
                if line.startswith('Ft site(0-3)') and line1[3].startswith('err'):
                    err = list_f[i].strip().split("****")[1].strip().split('err:')[1].split('#')[0].strip()
                    dic.update({line1[0]:err})
                elif line.startswith('Ft site(0-3)') and line1[4] == '#GL6104':
                    dic.update({line1[0]:''})
            else:
                dic.update({line1[0]:''})



    ###### V06
            if line.startswith('UHS1 scan'):
                dic.update({line1[0]+'_scan':line1[[n.startswith('scan') for n in line1].index(True)]})
            
    ######
            if line.startswith('rw_'):
                rw_lst.append(line1[5]+line1[6])
                if line.startswith('rw_0'):
                    dic.update({line1[0]+'_rw':rw_lst})
                    rw_lst = []

            if line.startswith('[Arduino] EOT'):  #0206
                end_time = list_f[i].split('****')[0].split(' ')[0][1:-1].strip()   #0206

            if line.startswith('FtBin'):
                ft_on = 0
                FTbin = line[5:].strip().split(',')
                #print(FTbin)

                if FTbin[0] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[0].strip()
                    outdf.loc[count]['Start_time'] = site1_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time                    #0206
                    outdf.loc[count]['Site'] = '1'
                    outdf.loc[count]['Error'] = dic['DbgS0:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS0:_scan'] #V06
                    outdf.loc[count]['1_SMbus'] = site1_1_smbus #V0729
                    outdf.loc[count]['1_DUT'] = site1_1_dut #V0729
                    outdf.loc[count]['1_L1_count'] = site1_1_l1_count #V0729
                    outdf.loc[count]['2_SMbus'] = site1_2_smbus #V0729
                    outdf.loc[count]['2_DUT'] = site1_2_dut #V0729
                    outdf.loc[count]['2_L1_count'] = site1_2_l1_count #V0729
                    outdf.loc[count]['LotID'] = LotID #V0729
                    outdf.loc[count]['Runcard'] = Runcard #V0729
                    outdf.loc[count]['LoadBoard'] = LoadBoard #V0729

                    dic_rw_lst = dic['DbgS0:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break
                    
                    site1_1_smbus = ''
                    site1_1_dut = ''
                    site1_1_l1_count = ''
                    site1_2_smbus = ''
                    site1_2_dut = ''
                    site1_2_l1_count = ''


                if FTbin[1] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[1].strip()
                    outdf.loc[count]['Start_time'] = site2_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '2'
                    outdf.loc[count]['Error'] = dic['DbgS1:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS1:_scan'] #V06

                    outdf.loc[count]['1_SMbus'] = site2_1_smbus #V0729
                    outdf.loc[count]['1_DUT'] = site2_1_dut #V0729
                    outdf.loc[count]['1_L1_count'] = site2_1_l1_count #V0729
                    outdf.loc[count]['2_SMbus'] = site2_2_smbus #V0729
                    outdf.loc[count]['2_DUT'] = site2_2_dut #V0729
                    outdf.loc[count]['2_L1_count'] = site2_2_l1_count #V0729
                    outdf.loc[count]['LotID'] = LotID #V0729
                    outdf.loc[count]['Runcard'] = Runcard #V0729
                    outdf.loc[count]['LoadBoard'] = LoadBoard #V0729
                    
                    dic_rw_lst = dic['DbgS1:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        if dic_rw_lst[n] != rw_bin1[n]:
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                    site2_1_smbus = ''
                    site2_1_dut = ''
                    site2_1_l1_count = ''
                    site2_2_smbus = ''
                    site2_2_dut = ''
                    site2_2_l1_count = ''


                if FTbin[2] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[2].strip()
                    outdf.loc[count]['Start_time'] = site3_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '3'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06

                    outdf.loc[count]['1_SMbus'] = site3_1_smbus #V0729
                    outdf.loc[count]['1_DUT'] = site3_1_dut #V0729
                    outdf.loc[count]['1_L1_count'] = site3_1_l1_count #V0729
                    outdf.loc[count]['2_SMbus'] = site3_2_smbus #V0729
                    outdf.loc[count]['2_DUT'] = site3_2_dut #V0729
                    outdf.loc[count]['2_L1_count'] = site3_2_l1_count #V0729
                    outdf.loc[count]['LotID'] = LotID #V0729
                    outdf.loc[count]['Runcard'] = Runcard #V0729
                    outdf.loc[count]['LoadBoard'] = LoadBoard #V0729

                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break


                    site3_1_smbus = ''
                    site3_1_dut = ''
                    site3_1_l1_count = ''
                    site3_2_smbus = ''
                    site3_2_dut = ''
                    site3_2_l1_count = ''


                if FTbin[3] != 'ff':  #寫入outdf
                    count += 1
                    outdf.loc[count]=''
                    outdf.loc[count]['FTBin'] = FTbin[3].strip()
                    outdf.loc[count]['Start_time'] = site4_start_time   #0206
                    outdf.loc[count]['End_time'] = end_time    
                    outdf.loc[count]['Site'] = '4'
                    outdf.loc[count]['Error'] = dic['DbgS2:']
                    outdf.loc[count]['UHS1 scan'] = dic['DbgS2:_scan'] #V06

                    outdf.loc[count]['1_SMbus'] = site4_1_smbus #V0729
                    outdf.loc[count]['1_DUT'] = site4_1_dut #V0729
                    outdf.loc[count]['1_L1_count'] = site4_1_l1_count #V0729
                    outdf.loc[count]['2_SMbus'] = site4_2_smbus #V0729
                    outdf.loc[count]['2_DUT'] = site4_2_dut #V0729
                    outdf.loc[count]['2_L1_count'] = site4_2_l1_count #V0729
                    outdf.loc[count]['LotID'] = LotID #V0729
                    outdf.loc[count]['Runcard'] = Runcard #V0729
                    outdf.loc[count]['LoadBoard'] = LoadBoard #V0729

                    dic_rw_lst = dic['DbgS2:_rw']
                    for n in range(0,len(dic_rw_lst)-1):
                        #print(n)
                        if dic_rw_lst[n] != rw_bin1[n]:
                            #print(n)
                            outdf.loc[count]['rw_fail'] = dic_rw_lst[n].split(':')[0]
                            break

                    site4_1_smbus = ''
                    site4_1_dut = ''
                    site4_1_l1_count = ''
                    site4_2_smbus = ''
                    site4_2_dut = ''
                    site4_2_l1_count = ''


        f.close()
        outdf['File'] = file
        return outdf



alldf = DataFrame(columns =['SN', 'File','Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan','Start_time','End_time','1_SMbus','1_DUT','1_L1_count','2_SMbus','2_DUT','2_L1_count']) 
alldf.index = alldf['SN']

path = 'log/'
dirs = os.listdir(path)


print('請選擇FT程式:\n1. 7428 V05\n2. 7428 V06_beta\n3. 7449 V03\n4. 7494 FT3 V02\n5. 7428 V07\n6. 7494 FT3 V03\n7. 7428 FT3 v08_2')
version_select = input()

if version_select == '1':
    for n in dirs:    
        alldf = concat([alldf,log_summary_v05(path+n)],axis=0)

if version_select == '2':
    for n in dirs:    
        alldf = concat([alldf,log_summary_v06(path+n)],axis=0)

if version_select == '3':
    for n in dirs:    
        alldf = concat([alldf,log_summary_7449_v03(path+n)],axis=0)

if version_select == '4':
    for n in dirs:    
        alldf = concat([alldf,log_summary_7494_FT3_v02(path+n)],axis=0)

if version_select == '5':
    for n in dirs:    
        alldf = concat([alldf,log_summary_7428_v07(path+n)],axis=0)

if version_select == '6':
    for n in dirs:    
        alldf = concat([alldf,log_summary_7494_FT3_v03(path+n)],axis=0)

if version_select == '7':
    for n in dirs:    
        alldf = concat([alldf,log_summary_7428_FT3_v08_2(path+n)],axis=0)

#alldf = alldf.reindex(columns = ['SN', 'File', 'Site', 'FTBin', 'Error', 'rw_fail', 'UHS1 scan'] )

shutil.copy('empty\\list.xlsx','list.xlsx')
book = load_workbook('list.xlsx')
writer = ExcelWriter('list.xlsx', engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
##            dfs.to_excel(writer, sheet_name='lot_raw data')
alldf.to_excel(writer, sheet_name='sheet1')

writer.save()
