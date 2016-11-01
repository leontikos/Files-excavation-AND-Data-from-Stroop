# imports
import numpy as np
import pandas as pd
import os
import glob
import shutil


# choosing the folder path
PTH = r"C:\Wytyk_ALL"
os.chdir(PTH)
print(os.getcwd())


# listing all important Excel Worksheets
sheets = ['mood2', 'slowa1', 'slowa2', 'picture_emo_seria1', 'emo_mix', 'emo_seria2']


# create an empty DB
db = pd.DataFrame(columns = ['pseudonim', 'imie'])
dt = pd.DataFrame()


# MAIN CODE FOR DATA EXCAVATION

for root, dirs, files in os.walk(PTH):
    for file in files:
        
        # reading Excel file
        dt1 = pd.ExcelFile(file)

        # walk through sheets
        for sheet in sheets:
            df = pd.read_excel(dt1, sheet)

            # walk through sheets and define variables accordingly
            if sheet == 'mood2':
                cols1 = ['pic', 'mood_2_rating.response_mean']
                cols2 = ['pic', 'mood_2_rating.rt_mean']
                df1 = df.loc[0:7, cols1]
                df2 = df.loc[0:7, cols2]

            elif sheet == 'slowa1':
                cols1 = ['congruent', 'resp_4.corr_mean']
                cols2 = ['congruent', 'resp_4.rt_mean']
                df1 = df.loc[0:66, cols1]
                df2 = df.loc[0:66, cols2]

            elif sheet == 'slowa2':
                cols1 = ['congruent', 'resp_4.corr_mean']
                cols2 = ['congruent', 'resp_4.rt_mean']
                df1 = df.loc[0:49, cols1]
                df2 = df.loc[0:49, cols2]   

            elif sheet == 'picture_emo_seria1':
                cols1 = ['Congruent', 'resp.corr_mean', 'word']
                cols2 = ['Congruent', 'resp.rt_mean', 'word']
                df1 = df.loc[0:64, cols1]
                df2 = df.loc[0:64, cols2] 

                # adding a letter to 'congruency' based on the 'word' column
                for i in df1.index:
                    if df1.loc[i, 'word'] == 'strach':
                        df1.loc[i, 'Congruent'] = df1.loc[i, 'Congruent'] + '_s'
                    else:
                        df1.loc[i, 'Congruent'] = df1.loc[i, 'Congruent'] + '_r'

                for i in df2.index:
                    if df2.loc[i, 'word'] == 'strach':
                        df2.loc[i, 'Congruent'] = df2.loc[i, 'Congruent'] + '_s'
                    else:
                        df2.loc[i, 'Congruent'] = df2.loc[i, 'Congruent'] + '_r'
                # dropping unnecessary column 'word'
                df1.drop('word', axis=1, inplace=True)
                df2.drop('word', axis=1, inplace=True)  

            elif sheet == 'emo_mix':
                cols1 = ['congruent', 'resp_4.corr_mean']
                cols2 = ['congruent', 'resp_4.rt_mean']
                df1 = df.loc[0:36, cols1]
                df2 = df.loc[0:36, cols2]    

            elif sheet == 'emo_seria2':    
                cols1 = ['Congruent', 'resp.corr_mean', 'word']
                cols2 = ['Congruent', 'resp.rt_mean', 'word']
                df1 = df.loc[0:36, cols1]
                df2 = df.loc[0:36, cols2] 

                # adding a letter to congruency based on the 'word' column
                for i in df1.index:
                    if df1.loc[i, 'word'] == 'strach':
                        df1.loc[i, 'Congruent'] = df1.loc[i, 'Congruent'] + '_s'
                    else:
                        df1.loc[i, 'Congruent'] = df1.loc[i, 'Congruent'] + '_r'

                for i in df2.index:
                    if df2.loc[i, 'word'] == 'strach':
                        df2.loc[i, 'Congruent'] = df2.loc[i, 'Congruent'] + '_s'
                    else:
                        df2.loc[i, 'Congruent'] = df2.loc[i, 'Congruent'] + '_r'

                # dropping unnecessary column 'word'
                df1.drop('word', axis=1, inplace=True)
                df2.drop('word', axis=1, inplace=True)

            # change column names in sheets
            df1.columns = ['pseudonim', 'imie']
            df2.columns = ['pseudonim', 'imie']

            # append the DB
            db = db.append(df1, ignore_index=True)
            db = db.append(df2, ignore_index=True)
        #    Activate for debug only:
        #    print('DONE: ' + sheet)

        # adding demographic data
        # loading 'mood2' worksheet
        plwiek = pd.read_excel(dt1, 'mood2')
        # adding columns to the demographics DB
        cols = ['nr', 'pic']
        plwiek_db = plwiek.loc[9:15, cols]
        # changing col names to prepare them for append
        plwiek_db.columns = ['pseudonim', 'imie']
        # appending
        db = db.append(plwiek_db, ignore_index=True)
        #print('DONE: demographics')

        # changing the name of the participant,
        # printing DONE file
        namecol = file.split('_')[0]
        db.rename(columns={db.columns[1]: namecol}, inplace=True)
        print('DONE: ' + namecol)

        if dt.empty:
            dt = db
        else:
            db.drop('pseudonim', axis=1, inplace=True)
            dt = dt.join(db)

        # cleaning DB in the end
        db = pd.DataFrame(columns = ['pseudonim', 'imie'])
print('***DONE***')

        # IN ORDER TO SAVE THE DATABASE SAVE 'DT' DATAFRAME TO EXCEL
				
# Saving the database
dt.to_excel('Wytykowska_Baza_Danych.xlsx')
