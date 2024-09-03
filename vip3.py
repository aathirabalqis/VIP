import os
import math
import shutil
import numpy as np
import pandas as pd 
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
from openpyxl import load_workbook

root = Tk()

root.title('VIP')
root.geometry('500x250') 

def getvip():
    global dfvip
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 180, y  = 32)
    
    dfvip = pd.read_excel(filenames[0])
    return dfvip
    
def getpast():
    global dfpast, filenames, pastpath
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 180, y  = 72)
    
    pastpath = filenames[0]
    dfpast = pd.read_excel(pastpath)
    return dfpast
    
def getsitrep():
    global sitreps
    sitreps = fd.askopenfilenames(filetypes=[("Text files","*.txt")])
    Label(root, text = sitreps).place(x = 180, y  = 112)
        
    return sitreps
    
def getinfo():
    global dfpast, sitreps, filenames, folder, name, date
    
    folder = '\\'.join(filenames[0].split('/')[:-1])
    name = filenames[0].split('/')[-1].split(' ')
    
    for i in sitreps:
        if '_' in i: date = i.split('/')[-1].split('_')[2]
    joined = folder + '\\' + ' '.join(name[:-3]) + ' ' + date + '.xlsx'
   
    
    # get vip list no tagging
    dfnotag = dfpast.loc[(dfpast['Star Customer Type'] == 'No'), ['Contract Account','Star Customer Type']]
    # dfnotag = dfnotag.rename(columns = {'Contract Acc':'Contract Account'})
    df = dfvip.append(dfnotag)
    
    for i in sitreps:
        q = open('temp.txt', 'w')
        
        # reformat sitrep files
        with open(i) as f: 
            for line in f:
                if '|' in line and 'Description' not in line and '----' not in line:
                    q.write(line)
                    
        q.close()
        print('done')
        
        # read new sitrep file
        dfsit = pd.read_csv('temp.txt', names = ['Empty','State','Station','Station Description','Installation','Contract Account','Telephone No.','B. Partner','Customer Name','Address','GPS Coordinate','Voltage Level','Rate Category','Installation Type','Logical device no.','Device No.','Device Cat.','Register Group','Meter Installation Date','DAT','Controlling Device','Portion','MR Unit','MRU Description','IP No.','AMS','AMCG','Landlord/Tenant','Installation No Landlord/Tenant'], index_col = False,sep = '|', encoding = "ISO-8859-1")
        dfsit = dfsit.iloc[:,1:]
        
        # merge with VIP CAs
        df = pd.merge(left = df, right = dfsit, how = 'left', on = 'Contract Account')
        
        for i in df.columns:
            if '_x' in i:
                y = i[:-2] + '_y'
                df.loc[(pd.isnull(df[i])), i] = df[y]
                
                df = df.rename(columns = {i:i[:-2]})
                df = df.drop(columns = [y])

    df = df.loc[(pd.isnull(df['State']) == False)]
    df['Count'] = df.groupby('State')['State'].transform(len)
    df.sort_values(['Count','State','Station'], ascending = [False, True, True], inplace = True)
    df = df[['State','Station','Station Description','Installation','Contract Account','Customer Name','Address','GPS Coordinate','Voltage Level','Rate Category','Installation Type','Device No.','Device Cat.','Meter Installation Date','Portion','MR Unit','MRU Description','AMS','AMCG','Star Customer Type']]
    # df.to_excel(joined, sheet_name = 'TOTAL (' + str(len(df)) + ')', index = False) 

    print('done')

# def export():
    # global filenames, sitreps, pastpath, dfpast
    
    # folder = '\\'.join(filenames[0].split('/')[:-1])
    # name = filenames[0].split('/')[-1].split(' ')
    # for i in sitreps:
        # if '_' in i: date = i.split('/')[-1].split('_')[2]
    # joined = folder + '\\' + ' '.join(name[:-3]) + ' ' + date + '.xlsx'
    # print(os.listdir(folder))
    # print(joined)
    
    # df = pd.read_excel(joined)
    # print('start')
    # print(df.columns)
    #
    
    # create reports to export
    df['State'] = df['State'].str.strip() 
    
    writer = pd.ExcelWriter(joined, engine = 'xlsxwriter')
    df.to_excel(writer, sheet_name = 'TOTAL (' + str(len(df)) + ')', index = False)
    
    df['Installation Type'] = df['Installation Type'].astype(int)
    df['AMCG'] = df['AMCG'].astype(int)
    
    print(df)
    
    # get nems sheet
    nems = df.loc[(df['Installation Type'].astype(str) == '25')]
    print(nems)
    nems.to_excel(writer, sheet_name = 'NEM (' + str(len(nems)) + ')', index = False)
    
    # get spot sheet
    spot = df.loc[(df['Portion'].str.startswith('SPOT')) & (df['AMCG'].astype(str) == '301')]
    print(spot)
    spot.to_excel(writer, sheet_name = 'SPOT (301)', index = False)
    
    # get vip sheet
    vip = df.loc[(df['AMCG'].astype(str) == '998')]
    print(vip)
    vip.to_excel(writer, sheet_name = 'VIP (998)', index = False)
    
    print('donez')
    
    # get states sheet
    kv = df.loc[(df['State'] == 'SEL') | (df['State'] == 'KUL') | (df['State'] == 'PJY')]
    kv = kv.loc[(kv['Installation Type'].astype(str) != '25')]
    kv.to_excel(writer, sheet_name = 'KV (' + str(len(kv)) + ')', index = False)
    
    mel = df.loc[(df['State'] == 'MEL') & (df['Installation Type'].astype(str) != '25')]
    mel.to_excel(writer, sheet_name = 'MEL (' + str(len(mel)) + ')', index = False)
    
    ked = df.loc[(df['State'] == 'KED') & (df['Installation Type'].astype(str) != '25')]
    ked.to_excel(writer, sheet_name = 'KED (' + str(len(ked)) + ')', index = False)
    
    oth = df.loc[(df['State'] != 'SEL') & (df['State'] != 'KUL') & (df['State'] != 'PJY') & (df['State'] != 'MEL') & (df['State'] != 'KED')]
    oth = oth.loc[(df['Installation Type'].astype(str) != '25')]
    oth.to_excel(writer, sheet_name = 'OTH (' + str(len(oth)) + ')', index = False)
    
    print('testz')
    writer.close()
    
    # make file without customer name
    newf = folder + '\\' + 'STATUS_LIST_SM_VVIP_' + date + '.xlsx'
    writer = pd.ExcelWriter(joined, mode = 'a', engine = 'openpyxl')
    writer2 = pd.ExcelWriter(newf, engine = 'xlsxwriter')
    
    for i in writer.book.sheetnames:
        tempdf = pd.read_excel(joined, sheet_name = i)
        tempdf = tempdf.drop(columns = ['Customer Name'])
        tempdf.to_excel(writer2, sheet_name = i, index = False)
    
    # make new report to include past analysis
    pastdf = pd.read_excel(pastpath, sheet_name = 'ANALYSIS KV')
    pastdf2 = pd.read_excel(pastpath, sheet_name = 'ANALYSIS NOT KV')
    pastdf3 = pd.read_excel(pastpath, sheet_name = 'VIP NOT SM')
    
    pastdf.to_excel(writer, sheet_name = 'ANALYSIS KV', index = False)
    pastdf2.to_excel(writer, sheet_name = 'ANALYSIS NOT KV', index = False)
    pastdf3.to_excel(writer, sheet_name = 'VIP NOT SM', index = False)
    
    # make new sheet of added and removed CAs
    print(df.columns)
    vipcur = df.loc[(df['Star Customer Type'] == 'VIP'),['Device No.','Contract Account','Customer Name']]
    vipcur['Status'] = 'Added'
    print(dfpast.columns)
    vippast = dfpast.loc[(dfpast['Star Customer Type'] == 'VIP'),['Device No.','Contract Account','Customer Name']]
    vippast['Status'] = 'Removed'
    
    inter = pd.merge(left = vipcur, right = vippast, how = 'inner', on = 'Contract Account')
    inter = inter['Contract Account'].tolist()
    
    newcur = vipcur.loc[(vipcur['Contract Account'].isin(inter) == False)]
    newpast = vippast.loc[(vippast['Contract Account'].isin(inter) == False)]
    
    lastvip = newcur.append(newpast)
    lastvip.to_excel(writer, sheet_name = 'VIP Changes', index = False)
    
    writer.close()
    writer2.close()

def test():
    print('yes')
      
def shortcut(event):
    if event.char == 'h':
        messagebox.showinfo(title=None, message='hakeem loser')

def exitt(event):
    root.quit()

    
Button(root, text = 'VIP List (BCRM)', command = getvip).place(x = 38, y = 30) 
Button(root, text = 'Past VIP Report', command = getpast).place(x = 40, y = 70)
Button(root, text = 'SITREP (BCRM)', command = getsitrep).place(x = 41, y = 110)

Label(root, text = 'Path:').place(x = 140, y  = 32) 
Label(root, text = 'Path:').place(x = 140, y  = 72) 
Label(root, text = 'Path:').place(x = 140, y  = 112) 

Button(root, text = 'Extract Information and Export for Dashboard & Monitoring', command = getinfo).place(x = 90, y = 170)
# Button(root, text = 'Export for Dashboard & Monitoring', command = export).place(x = 205, y = 170)

Label(root, text = 'SAB').place(x = 470, y = 230)

root.resizable(True, False) 
root.bind('<Key>', shortcut)
root.bind('<Escape>', exitt)
root.mainloop()