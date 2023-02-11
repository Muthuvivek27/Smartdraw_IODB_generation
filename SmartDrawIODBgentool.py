import pandas as pd
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as mb
from tkinter import ttk
import os
import numpy as np
import win32com.client as win32
from sys import exit
import xlsxwriter


#Export Mapping CSV for Further Use:
def CreateCSV():
    global df3, Filename,csv
    df3 = pd.DataFrame({"Def_map":Def_map, "Mapping_Field":map})
    csv=Tk()
    csv.geometry('300x300')
    Filename = Entry(csv)
    l = ttk.Label(csv,text="Filename").grid(column=0, row=0)
    Filename.grid(row=1, column=0)
    button1 = Button(csv, text="SAVE", command=save, width=10, state="normal").grid(row=2, column=0)

def save():
    filename = Filename.get()
    csv_filename = csv_path+'\\Mapping_' + filename + '.csv'
    df3.to_csv(csv_filename, index=False)
    csv.destroy()
    mb.showinfo("INFO","Your Mapping File has been Exported to"+ csv_filename)


#import CSV for mapping
def ImportCSV():
    global dmap
    print(csv_path)
    prime_field=[]
    map=[]
    file_map = filedialog.askopenfilename(initialdir= csv_path, title="Mapping File",
                                           filetypes=(("CSV", "*.csv"), ("all files", "*.*")))
    if file_map=='':
        mb.showerror("Error","Mapping File Not Selected properly:")
        master.destroy()
        mapping()
    dmap = pd.read_csv(file_map, index_col=False).fillna('')
    prime_field = dmap['Def_map'].tolist()
    map = dmap['Mapping_Field'].tolist()
    mb.showinfo('Info','Mapping file has been Imported.')
    check(map)
    IODB(prime_field,map)

def rev_update():
    global df_IODB_NEW_sys, df_IODB_NEW_Node, df_IODB_NEW_Board, df_IODB_NEW_TB, df_IODB_NEW_barrier
    global df_IODB_OLD_sys, df_IODB_OLD_Node, df_IODB_OLD_Board, df_IODB_OLD_TB, df_IODB_OLD_barrier

    #system
    df_IODB_NEW_sys = df_IODB_NEW_sys.merge(df_IODB_OLD_sys, indicator=True, how='outer')

    df_temp = df_IODB_NEW_sys

    conditions = [
        (df_temp['_merge'] == 'both'),
        (df_temp['_merge'] == 'left_only'),
        (df_temp['_merge'] == 'right_only')]
    choices = ['', rev_number, 'DEL']

    #getting rev updates in both old and new columns
    df_temp['Rev_Slot'] = np.select(conditions, choices, default='null')

    #dropping '_merge' so that we can hav euniform columns
    df_temp = df_temp.drop(['_merge'], axis=1)

    #finding repetitive items in the group
    df_IODB_new_1 = df_temp.groupby(['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT']).size().reset_index(
        name='total')

    #filtering the modified data alone.
    df_IODB_new_2 = df_IODB_new_1[df_IODB_new_1.total > 1]

    #getting right only for comman as both
    df_IODB_new_3 = df_temp.merge(df_IODB_new_2, indicator=True, how='outer')

    #dropping our needed case
    df_IODB_new_4 = df_IODB_new_3[(df_IODB_new_3['_merge'] != 'both') | (df_IODB_new_3['Rev_Slot'] != 'DEL')]

    #final-system
    df_IODB_NEW_sys = df_IODB_new_4.drop(['_merge', 'total'], axis=1).sort_values(
        by=['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT', 'REDUNDANCY_SLOT', 'IO_MODULE']).fillna("")





    #Node
    df_IODB_NEW_Node = df_IODB_NEW_Node.merge(df_IODB_OLD_Node, indicator=True, how='outer')

    df_temp = df_IODB_NEW_Node

    conditions = [
        (df_temp['_merge'] == 'both'),
        (df_temp['_merge'] == 'left_only'),
        (df_temp['_merge'] == 'right_only')]
    choices = ['', rev_number, 'DEL']

    # getting rev updates in both old and new columns
    df_temp['Rev_Node'] = np.select(conditions, choices, default='null')

    # dropping '_merge' so that we can hav euniform columns
    df_temp = df_temp.drop(['_merge'], axis=1)

    # finding repetitive items in the group
    df_IODB_new_1 = df_temp.groupby(['SYSTEM_CABINET', 'NODE']).size().reset_index(
        name='total')

    # filtering the modified data alone.
    df_IODB_new_2 = df_IODB_new_1[df_IODB_new_1.total > 1]

    # getting right only for commOn as both
    df_IODB_new_3 = df_temp.merge(df_IODB_new_2, indicator=True, how='outer')

    # dropping our needed case
    df_IODB_new_4 = df_IODB_new_3[(df_IODB_new_3['_merge'] != 'both') | (df_IODB_new_3['Rev_Node'] != 'DEL')]

    # final
    df_IODB_NEW_Node = df_IODB_new_4.drop(['_merge', 'total'], axis=1).sort_values(
        by=['SYSTEM_CABINET', 'NODE']).fillna("")


    #Boards
    df_IODB_NEW_Board = df_IODB_NEW_Board.merge(df_IODB_OLD_Board, indicator=True, how='outer')
    df_temp = df_IODB_NEW_Board

    conditions = [
        (df_temp['_merge'] == 'both'),
        (df_temp['_merge'] == 'left_only'),
        (df_temp['_merge'] == 'right_only')]
    choices = ['', rev_number, 'DEL']

    # getting rev updates in both old and new columns
    df_temp['Rev_Board'] = np.select(conditions, choices, default='null')

    # dropping '_merge' so that we can hav euniform columns
    df_temp = df_temp.drop(['_merge'], axis=1)

    # finding repetitive items in the group
    df_IODB_new_1 = df_temp.groupby(['BOARD_IN_MPNAME', 'BOARD_NAME']).size().reset_index(
        name='total')

    # filtering the modified data alone.
    df_IODB_new_2 = df_IODB_new_1[df_IODB_new_1.total > 1]

    # getting right only for comman as both
    df_IODB_new_3 = df_temp.merge(df_IODB_new_2, indicator=True, how='outer')

    # dropping our needed case
    df_IODB_new_4 = df_IODB_new_3[(df_IODB_new_3['_merge'] != 'both') | (df_IODB_new_3['Rev_Board'] != 'DEL')]

    # final
    df_IODB_NEW_Board = df_IODB_new_4.drop(['_merge', 'total'], axis=1).sort_values(
        by=['BOARD_IN_MPNAME', 'BOARD_NAME']).fillna("")



    # TB
    df_IODB_NEW_TB = df_IODB_NEW_TB.merge(df_IODB_OLD_TB, indicator=True, how='outer')

    df_temp = df_IODB_NEW_TB

    conditions = [
        (df_temp['_merge'] == 'both'),
        (df_temp['_merge'] == 'left_only'),
        (df_temp['_merge'] == 'right_only')]
    choices = ['', rev_number, 'DEL']

    # getting rev updates in both old and new columns
    df_temp['Rev_TB'] = np.select(conditions, choices, default='null')

    # dropping '_merge' so that we can hav euniform columns
    df_temp = df_temp.drop(['_merge'], axis=1)

    # finding repetitive items in the group
    df_IODB_new_1 = df_temp.groupby(['JBCABLE_IN_MPNAME', 'MP_TS_NAME']).size().reset_index(
        name='total')

    # filtering the modified data alone.
    df_IODB_new_2 = df_IODB_new_1[df_IODB_new_1.total > 1]

    # getting right only for comman as both
    df_IODB_new_3 = df_temp.merge(df_IODB_new_2, indicator=True, how='outer')

    # dropping our needed case
    df_IODB_new_4 = df_IODB_new_3[(df_IODB_new_3['_merge'] != 'both') | (df_IODB_new_3['Rev_TB'] != 'DEL')]

    # final
    df_IODB_NEW_TB = df_IODB_new_4.drop(['_merge', 'total'], axis=1).sort_values(
        by=['JBCABLE_IN_MPNAME', 'MP_TS_NAME']).fillna("")


    # barrier Count
    df_IODB_NEW_barrier = df_IODB_OLD_barrier.merge(df_IODB_NEW_barrier, indicator=True, how='left')
    df_IODB_NEW_barrier['REV_BARRIER'] = df_IODB_NEW_barrier['_merge'].apply(lambda Z5: rev_number if Z5 != 'both' else '')
    df_IODB_NEW_barrier = df_IODB_NEW_barrier.drop(['_merge'], axis=1)
    button2 = ttk.Button(exporttk, text="Export IODB", command=Export, state="normal", width=15).grid(row=2, column=0,
                                                                                                      padx=40)
    exporttk.update()

def Save_IODB():
    global df_Cabinet_name, df_Alarm, df_PDP

    final_file_name = IodbFilename.get()
    export_file_name = excel_path + "\\" + str(final_file_name) + '.xlsx'
    export_file.destroy()
    writer = pd.ExcelWriter(export_file_name, engine='xlsxwriter')

    df_Cabinet_name.to_excel(writer, sheet_name='CabinetName', index=False)
    df_Alarm.to_excel(writer, sheet_name='Alarm', index=False)
    df_PDP.to_excel(writer, sheet_name='PDP', index=False)
    df_IODB_NEW.to_excel(writer, sheet_name='IODB', index=False)
    df_IODB_NEW_sys.to_excel(writer, sheet_name='System', index=False)
    df_IODB_NEW_Node.to_excel(writer, sheet_name='Node', index=False)
    df_IODB_NEW_Board.to_excel(writer, sheet_name='Board', index=False)
    df_IODB_NEW_TB.to_excel(writer, sheet_name='TerminalStrips', index=False)
    df_IODB_NEW_barrier.to_excel(writer, sheet_name='Barrier', index=False)
    df_IODB_NEW_Isolator.to_excel(writer, sheet_name='Isolator', index=False)
    df_IODB_NEW_relay.to_excel(writer, sheet_name='Relay', index=False)
    df_IODB_NEW_IRP_TS.to_excel(writer, sheet_name='TerminalStrips_IRP', index=False)
    df_IODB_NEW_IRP_MCC_TS.to_excel(writer, sheet_name='TerminalStrips-MCC', index=False)

    writer.save()

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(export_file_name)
    ws1 = wb.Worksheets("CabinetName").Columns.AutoFit()
    ws2 = wb.Worksheets("Alarm").Columns.AutoFit()
    ws3 = wb.Worksheets("PDP").Columns.AutoFit()
    ws4 = wb.Worksheets("IODB").Columns.AutoFit()
    ws5 = wb.Worksheets("System").Columns.AutoFit()
    ws6 = wb.Worksheets("Node").Columns.AutoFit()
    ws7 = wb.Worksheets("Board").Columns.AutoFit()
    ws8 = wb.Worksheets("TerminalStrips").Columns.AutoFit()
    ws9 = wb.Worksheets("Barrier").Columns.AutoFit()
    ws10 = wb.Worksheets("Isolator").Columns.AutoFit()
    ws11 = wb.Worksheets("Relay").Columns.AutoFit()
    ws12 = wb.Worksheets("TerminalStrips_IRP").Columns.AutoFit()
    ws13 = wb.Worksheets("TerminalStrips-MCC").Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


    mb.showinfo("Info", "Your File have been Exported to:"+ excel_path)


def Export():

    global IodbFilename
    global export_file
    export_file =Tk()
    export_file.title('FileName')
    export_file.geometry('200x200')
    IodbFilename = Entry(export_file)
    l = ttk.Label(export_file, text="Filename").grid(column=0, row=0)
    IodbFilename.grid(row=1, column=0)

    button1 = ttk.Button(export_file, text="SAVE", command=Save_IODB, width=10, state="normal").grid(row=2, column=0)


def IODB(Default_map,map_final):

    global df_IODB_NEW, df_IODB_OLD
    prime_field=Default_map
    map1=map_final
    lenth = len(map1)
    df_IODB_NEW = pd.DataFrame()
    df_IODB_OLD = pd.DataFrame()
    e=[]
    err='ok'
    for i in range(0, lenth):
        try:
            if map_final[i] == '':
                df_IODB_NEW[prime_field[i]]=""
                if x == 1:
                    df_IODB_OLD[prime_field[i]]=""
            else:
                df_IODB_NEW[prime_field[i]] = df[map_final[i]]
                if x==1:
                    df_IODB_OLD[prime_field[i]] = df1[map_final[i]]
        except KeyError as e:
            if str(e)!='':
                mb.showerror("Error","Header Column " + str(e) + " Not Found!Please check:")
                err='not ok'
    if err == 'not ok':
        mb.showerror("Error","Please Check Mapping File for Headers not available in IODB.")
        master.destroy()
        mapping()


    map_status = 1
    if map_status == 1:
        button2 = Button(master, text="SAVE", command=CreateCSV, state="normal", width=10).grid(row=19, column=4)
        button4 = Button(master, text="CONTINUE", command=Conti, state="normal", width=25,bg='#82E0AA').grid(row=21, column=4)
    master.update()


def Pivot_sheet():
    global df_IODB_NEW_sys, df_IODB_NEW_Node,df_IODB_NEW_Board,df_IODB_NEW_TB,df_IODB_NEW_barrier,df_IODB_NEW_Isolator
    global df_IODB_NEW_relay,df_IODB_NEW_IRP_TS,df_IODB_NEW_IRP_MCC_TS
    global df_IODB_OLD_sys, df_IODB_OLD_Node, df_IODB_OLD_Board, df_IODB_OLD_TB, df_IODB_OLD_barrier

    #For all system details:
    df_IODB_NEW_sys = df_IODB_NEW[['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT', 'REDUNDANCY_SLOT',
                                     'IO_MODULE']].drop_duplicates().sort_values(
        by=['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT', 'REDUNDANCY_SLOT', 'IO_MODULE']).fillna("")

    #For Node
    df_IODB_NEW_Node = df_IODB_NEW[['SYSTEM_CABINET', 'NODE']].drop_duplicates().sort_values(
        by=['SYSTEM_CABINET', 'NODE']).fillna("")
    # For Boards
    df_IODB_NEW_Board = df_IODB_NEW[['BOARD_IN_MPNAME', 'BOARD_NAME','SYSTEM_CABINET','NODE','SLOT','REDUNDANCY_SLOT', 'BOARD_MODEL', 'IO_TYPE']].drop_duplicates().sort_values(
        by=['BOARD_IN_MPNAME', 'BOARD_NAME', 'BOARD_MODEL', 'IO_TYPE']).fillna("")

    # For TB
    df_IODB_NEW_TB = df_IODB_NEW[
        ['JBCABLE_IN_MPNAME', 'MP_TS_NAME', 'JB_CABLE_NM', 'JB_CABLE_TYPE']].drop_duplicates().sort_values(
        by=['JBCABLE_IN_MPNAME', 'MP_TS_NAME', 'JB_CABLE_NM', 'JB_CABLE_TYPE']).fillna("")

    #For Barrier Count
    df_temp = df_IODB_NEW[['BOARD_IN_MPNAME','BOARD_NAME', 'BARRIER_MODEL']].fillna("")
    df_temp = df_temp[df_temp['BARRIER_MODEL'] != '']
    df_IODB_NEW_barrier = df_temp.groupby(['BOARD_IN_MPNAME','BOARD_NAME', 'BARRIER_MODEL']).size().reset_index(name='total')


    #Isolator
    df_IODB_NEW_Isolator = df_IODB_NEW[
        ['BOARD_IN_MPNAME', 'BOARD_ISOLATOR_MODEL', 'BOARD_ISOLATOR_NAME']].drop_duplicates().sort_values(
        by=['BOARD_IN_MPNAME', 'BOARD_ISOLATOR_MODEL', 'BOARD_ISOLATOR_NAME']).fillna("")

    #Relay
    df_IODB_NEW_relay = df_IODB_NEW[
        ['RLY_IN_MP_NAME', 'RLY_NAME', 'RLY_MODEL']].drop_duplicates().sort_values(
        by=['RLY_IN_MP_NAME', 'RLY_NAME', 'RLY_MODEL']).fillna("")

    # IRP_TS
    df_IODB_NEW_IRP_TS = df_IODB_NEW[
        ['IRP_MP_NAME', 'IRP_MAR_TS_NM', 'IRP_MAR_CABLE_NM','IRP_MAR_CABLE_TYPE']].drop_duplicates().sort_values(
        by=['IRP_MP_NAME', 'IRP_MAR_TS_NM', 'IRP_MAR_CABLE_NM','IRP_MAR_CABLE_TYPE']).fillna("")

    #IRP_MCC_TS
    df_IODB_NEW_IRP_MCC_TS = df_IODB_NEW[
        ['IRP_MP_NAME', 'IRP_MCC_TS_NM', 'IRP_MCC_CABLE_NM', 'IRP_MCC_CABLE_TYPE']].drop_duplicates().sort_values(
        by=['IRP_MP_NAME', 'IRP_MCC_TS_NM', 'IRP_MCC_CABLE_NM', 'IRP_MCC_CABLE_TYPE']).fillna("")

    if x ==1:
        # For all system details:
        df_IODB_OLD_sys = df_IODB_OLD[['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT', 'REDUNDANCY_SLOT',
                                       'IO_MODULE']].drop_duplicates().sort_values(
            by=['SYSTEM_CABINET', 'CONTROLLER_NAME', 'NODE', 'SLOT', 'REDUNDANCY_SLOT', 'IO_MODULE']).fillna("")

        # For Node
        df_IODB_OLD_Node = df_IODB_OLD[['SYSTEM_CABINET', 'NODE']].drop_duplicates().sort_values(
            by=['SYSTEM_CABINET', 'NODE']).fillna("")
        # For Boards
        df_IODB_OLD_Board = df_IODB_OLD[
            ['BOARD_IN_MPNAME', 'BOARD_NAME','SYSTEM_CABINET','NODE','SLOT','REDUNDANCY_SLOT', 'BOARD_MODEL', 'IO_TYPE']].drop_duplicates().sort_values(
            by=['BOARD_IN_MPNAME', 'BOARD_NAME', 'BOARD_MODEL', 'IO_TYPE']).fillna("")

        # For TB
        df_IODB_OLD_TB = df_IODB_OLD[
            ['JBCABLE_IN_MPNAME', 'MP_TS_NAME', 'JB_CABLE_NM', 'JB_CABLE_TYPE']].drop_duplicates().sort_values(
            by=['JBCABLE_IN_MPNAME', 'MP_TS_NAME', 'JB_CABLE_NM', 'JB_CABLE_TYPE']).fillna("")

        # For Barrier Count
        df_temp = df_IODB_OLD[['BOARD_IN_MPNAME','BOARD_NAME', 'BARRIER_MODEL']].fillna("")
        df_temp = df_temp[df_temp['BARRIER_MODEL'] != '']

        df_IODB_OLD_barrier = df_temp.groupby(['BOARD_IN_MPNAME','BOARD_NAME', 'BARRIER_MODEL']).size().reset_index(name='total')

        rev_update()
    button2 = Button(exporttk, text="Export IODB", command=Export, state="normal", width=15).grid(row=2, column=0,
                                                                                                        padx=40)
    exporttk.update()
    mb.showinfo('Info','Pivot Sheet has been Generated.')

#navigation to Export frame
def Conti():
    master.destroy()
    global exporttk
    exporttk = Tk()
    exporttk.title("Pivot And Export")
    exporttk.geometry('400x600')
    exporttk.configure(bg='#E8DAEF')
    l = Label(exporttk, text="         EXPORT",bg='#E8DAEF').grid(column=0, row=0, columnspan=3,  pady=40)
    button1 = Button(exporttk, text="Pivot Sheet", width=15, state="normal",command=Pivot_sheet).grid(row=1, column=0,
                                                                                                       padx=40)
    button2 = Button(exporttk, text="Export IODB", command=Export, state="disabled", width=15).grid(row=2, column=0,
                                                                                                    padx=40)
    button3 = Button(exporttk, text="Home", command=homescreen, state="normal", width=15).grid(row=4, column=0)
    button4 = Button(exporttk, text="Mapping", command=mapping, state="normal", width=15).grid(row=4, column=1)
    button7 = Button(exporttk, text="Exit", command=exit, state="normal", width=15).grid(row=5, column=1, columnspan=2)
    l1 = Label(exporttk, text="\n\n\n\n\n\n",bg='#E8DAEF').grid(column=0, row=3, columnspan=2, padx=90, pady=40)

# compare Header
def compare():
    global headermatch
    global df
    global df1

    df1 = pd.read_excel(filename1,'IODB')
    df = pd.read_excel(filename2,'IODB')
    b = df.columns
    d = df1.columns

    len1 = len(b)
    len2 = len(d)

    if len1 == len2:
        c = sum(b == d)
        if c == len1:
            mb.showinfo("info", "Headers Matching")
            headermatch=1
        else:
            mb.showerror("Error", "Please Check headers")
            headermatch=0
            homescreen()
    else:
        mb.showerror("Error", "Please Check headers")
        headermatch=0
        homescreen()

# To get two inputs:
def file_get_2():
    root = Tk()
    global filename1
    global filename2
    mb.showinfo("OLD IODB","Please Select Previous Rev of IODB")
    filename1 = filedialog.askopenfilename(initialdir="/", title="Old Revision",
                                           filetypes=(("EXCEL", "*.xlsx"), ("all files", "*.*")))
    print(filename1)
    mb.showinfo("NEW IODB", "Please Select Latest Rev of IODB")
    filename2 = filedialog.askopenfilename(initialdir="/", title="New Revision",
                                           filetypes=(("EXCEL", "*.xlsx"), ("all files", "*.*")))
    print(filename2)
    root.destroy()

# To get single File
def file_get_1():
    root = Tk()
    global filename1
    filename1 = filedialog.askopenfilename(initialdir="/", title="IODB",
                                           filetypes=(("EXCEL", "*.xlsx"), ("all files", "*.*")))
    print(filename1)

    root.destroy()


def getmapping():
    global map

    map = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
           '', '', '', '', '']
    map[0] = PRI_FIELD_1.get()
    map[1] = PRI_FIELD_2.get()
    map[2] = PRI_FIELD_3.get()
    map[3] = PRI_FIELD_4.get()
    map[4] = PRI_FIELD_5.get()
    map[5] = PRI_FIELD_6.get()
    map[6] = PRI_FIELD_7.get()
    map[7] = PRI_FIELD_8.get()
    map[8] = PRI_FIELD_9.get()
    map[9] = PRI_FIELD_10.get()
    map[10] = PRI_FIELD_11.get()
    map[11] = PRI_FIELD_12.get()
    map[12] = PRI_FIELD_13.get()
    map[13] = PRI_FIELD_14.get()
    map[14] = PRI_FIELD_15.get()
    map[15] = PRI_FIELD_16.get()
    map[16] = PRI_FIELD_17.get()
    map[17] = PRI_FIELD_18.get()
    map[18] = PRI_FIELD_19.get()
    map[19] = PRI_FIELD_20.get()
    map[20] = PRI_FIELD_21.get()
    map[21] = PRI_FIELD_22.get()
    map[22] = PRI_FIELD_23.get()
    map[23] = PRI_FIELD_24.get()
    map[24] = PRI_FIELD_25.get()
    map[25] = PRI_FIELD_26.get()
    map[26] = PRI_FIELD_27.get()
    map[27] = PRI_FIELD_28.get()
    map[28] = PRI_FIELD_29.get()
    map[29] = PRI_FIELD_30.get()
    map[30] = PRI_FIELD_31.get()
    map[31] = PRI_FIELD_32.get()
    map[32] = PRI_FIELD_33.get()

    check(map)
    if check_map=='good':
        IODB(Def_map,map)
    map_status==1
    if map_status==1:
        button2 = Button(master, text="SAVE", command=CreateCSV, state="normal", width=10).grid(row=19, column=4)
        button4 = Button(master, text="CONTINUE", command=Conti, state="normal", width=25,bg='#82E0AA').grid(row=21, column=4)
    master.update()

def check(map1):
    global check_map
    check_map = 'good'
    print(map1)
    c = 'ok'
    for i in range(len(map1)):
       if map1[i] == '':
           c = 'notok'

    #if we want to compare duplication also enable the next line:
    #if len(map1) != len(set(map1)) or c =='notok':
    if c =='notok':
        check_map = 'error'
        mb.showerror("Warning",'Mapping Field Contains Blanks')
        #master.destroy()
        #mapping()



def mapping():
    global map_status
    global master
    global df
    global df1
    map_status=0
    global PRI_FIELD_1,PRI_FIELD_2,PRI_FIELD_3,PRI_FIELD_4,PRI_FIELD_5,PRI_FIELD_6,PRI_FIELD_7,PRI_FIELD_8,PRI_FIELD_9,PRI_FIELD_10,PRI_FIELD_11,PRI_FIELD_12,PRI_FIELD_13,PRI_FIELD_14,PRI_FIELD_15,PRI_FIELD_16,PRI_FIELD_17,PRI_FIELD_18,PRI_FIELD_19,PRI_FIELD_20,PRI_FIELD_21,PRI_FIELD_22,PRI_FIELD_23,PRI_FIELD_24,PRI_FIELD_25,PRI_FIELD_26,PRI_FIELD_27,PRI_FIELD_28,PRI_FIELD_29,PRI_FIELD_30,PRI_FIELD_31,PRI_FIELD_32,PRI_FIELD_33
    global Def_map
    try:
        exporttk.destroy()
    except:
        print("")

    Def_map = ['INSTRUMENT_TAG','UNIT_NO','SYSTEM_CABINET',
               'CONTROLLER_NAME','NODE','SLOT','REDUNDANCY_SLOT','IO_MODULE','BOARD_IN_MPNAME',
               'BOARD_MODEL','BOARD_NAME','BARRIER_MODEL','BARRIER_NAME','IO_TYPE',
               'JBCABLE_IN_MPNAME','JB_CABLE_NM','JB_CABLE_TYPE','MP_TS_NAME',
               'BOARD_ISOLATOR_MODEL','BOARD_ISOLATOR_NAME','RLY_IN_MP_NAME','RLY_MODEL',
               'RLY_NAME','IRP_MP_NAME','IRP_MAR_TS_NM','IRP_MAR_CABLE_NM','IRP_MAR_CABLE_TYPE','IRP_RELAY MODEL',
               'IRP_RELAY_NAME','IRP_MCC_TS_NM','IRP_MCC_CABLE_NM','IRP_MCC_CABLE_TYPE','EPC_REV_NO']
    options = df.columns.tolist()
    master = Tk()
    master.title("MAPPING")

    master.geometry('1000x1000')
    master.configure(bg='#FDEBD0')
    PRI_FIELD_1 = StringVar(master)
    PRI_FIELD_2 = StringVar(master)
    PRI_FIELD_3 = StringVar(master)
    PRI_FIELD_4 = StringVar(master)
    PRI_FIELD_5 = StringVar(master)
    PRI_FIELD_6 = StringVar(master)
    PRI_FIELD_7 = StringVar(master)
    PRI_FIELD_8 = StringVar(master)
    PRI_FIELD_9 = StringVar(master)
    PRI_FIELD_10 = StringVar(master)
    PRI_FIELD_11 = StringVar(master)
    PRI_FIELD_12 = StringVar(master)
    PRI_FIELD_13 = StringVar(master)
    PRI_FIELD_14 = StringVar(master)
    PRI_FIELD_15 = StringVar(master)
    PRI_FIELD_16 = StringVar(master)
    PRI_FIELD_17 = StringVar(master)
    PRI_FIELD_18 = StringVar(master)
    PRI_FIELD_19 = StringVar(master)
    PRI_FIELD_20 = StringVar(master)
    PRI_FIELD_21 = StringVar(master)
    PRI_FIELD_22 = StringVar(master)
    PRI_FIELD_23 = StringVar(master)
    PRI_FIELD_24 = StringVar(master)
    PRI_FIELD_25 = StringVar(master)
    PRI_FIELD_26 = StringVar(master)
    PRI_FIELD_27 = StringVar(master)
    PRI_FIELD_28 = StringVar(master)
    PRI_FIELD_29 = StringVar(master)
    PRI_FIELD_30 = StringVar(master)
    PRI_FIELD_31 = StringVar(master)
    PRI_FIELD_32 = StringVar(master)
    PRI_FIELD_33 = StringVar(master)
    l = Label(master,text="MAPPING\n",bg='#FDEBD0',font=("Courier", 12)).grid(column=0, row=0, columnspan=5)
    l1 = Label(master, text=Def_map[0],bg='#FDEBD0').grid(column=0, row=1)
    l2 = Label(master, text=Def_map[1],bg='#FDEBD0').grid(column=0, row=2)
    l3 = Label(master, text=Def_map[2],bg='#FDEBD0').grid(column=0, row=3)
    l4 = Label(master, text=Def_map[3],bg='#FDEBD0').grid(column=0, row=4)
    l5 = Label(master, text=Def_map[4],bg='#FDEBD0').grid(column=0, row=5)
    l6 = Label(master, text=Def_map[5],bg='#FDEBD0').grid(column=0, row=6)
    l7 = Label(master, text=Def_map[6],bg='#FDEBD0').grid(column=0, row=7)
    l8 = Label(master, text=Def_map[7],bg='#FDEBD0').grid(column=0, row=8)
    l9 = Label(master, text=Def_map[8],bg='#FDEBD0').grid(column=0, row=9)
    l10 = Label(master, text=Def_map[9],bg='#FDEBD0').grid(column=0, row=10)
    l11 = Label(master, text=Def_map[10],bg='#FDEBD0').grid(column=0, row=11)
    l12 = Label(master, text=Def_map[11],bg='#FDEBD0').grid(column=0, row=12)
    l13 = Label(master, text=Def_map[12],bg='#FDEBD0').grid(column=0, row=13)
    l14 = Label(master, text=Def_map[13],bg='#FDEBD0').grid(column=0, row=14)
    l15 = Label(master, text=Def_map[14],bg='#FDEBD0').grid(column=0, row=15)
    l16 = Label(master, text=Def_map[15],bg='#FDEBD0').grid(column=0, row=16)
    l17 = Label(master, text=Def_map[16],bg='#FDEBD0').grid(column=0, row=17)
    l18 = Label(master, text=Def_map[17],bg='#FDEBD0').grid(column=3, row=1)
    l19 = Label(master, text=Def_map[18],bg='#FDEBD0').grid(column=3, row=2)
    l20 = Label(master, text=Def_map[19],bg='#FDEBD0').grid(column=3, row=3)
    l21 = Label(master, text=Def_map[20],bg='#FDEBD0').grid(column=3, row=4)
    l22 = Label(master, text=Def_map[21],bg='#FDEBD0').grid(column=3, row=5)
    l23 = Label(master, text=Def_map[22],bg='#FDEBD0').grid(column=3, row=6)
    l24 = Label(master, text=Def_map[23],bg='#FDEBD0').grid(column=3, row=7)
    l25 = Label(master, text=Def_map[24],bg='#FDEBD0').grid(column=3, row=8)
    l26 = Label(master, text=Def_map[25],bg='#FDEBD0').grid(column=3, row=9)
    l27 = Label(master, text=Def_map[26],bg='#FDEBD0').grid(column=3, row=10)
    l28 = Label(master, text=Def_map[27],bg='#FDEBD0').grid(column=3, row=11)
    l29 = Label(master, text=Def_map[28],bg='#FDEBD0').grid(column=3, row=12)
    l30 = Label(master, text=Def_map[29],bg='#FDEBD0').grid(column=3, row=13)
    l31 = Label(master, text=Def_map[30],bg='#FDEBD0').grid(column=3, row=14)
    l32 = Label(master, text=Def_map[31],bg='#FDEBD0').grid(column=3, row=15)
    l33 = Label(master, text=Def_map[32],bg='#FDEBD0').grid(column=3, row=16)
    l34 = Label(master, text="IMPORT CSV",bg='#FDEBD0').grid(column=6, row=7)
    combobox1 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_1)
    combobox1['values'] = options
    combobox1.grid(column=1, row=1)
    combobox2 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_2)
    combobox2['values'] = options
    combobox2.grid(column=1, row=2)
    combobox3 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_3)
    combobox3['values'] = options
    combobox3.grid(column=1, row=3)
    combobox4 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_4)
    combobox4['values'] = options
    combobox4.grid(column=1, row=4)
    combobox5 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_5)
    combobox5['values'] = options
    combobox5.grid(column=1, row=5)
    combobox6 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_6)
    combobox6['values'] = options
    combobox6.grid(column=1, row=6)
    combobox7 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_7)
    combobox7['values'] = options
    combobox7.grid(column=1, row=7)
    combobox8 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_8)
    combobox8['values'] = options
    combobox8.grid(column=1, row=8)
    combobox9 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_9)
    combobox9['values'] = options
    combobox9.grid(column=1, row=9)
    combobox10 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_10)
    combobox10['values'] = options
    combobox10.grid(column=1, row=10)
    combobox11 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_11)
    combobox11['values'] = options
    combobox11.grid(column=1, row=11)
    combobox12 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_12)
    combobox12['values'] = options
    combobox12.grid(column=1, row=12)
    combobox13 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_13)
    combobox13['values'] = options
    combobox13.grid(column=1, row=13)
    combobox14 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_14)
    combobox14['values'] = options
    combobox14.grid(column=1, row=14)
    combobox15 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_15)
    combobox15['values'] = options
    combobox15.grid(column=1, row=15)
    combobox16 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_16)
    combobox16['values'] = options
    combobox16.grid(column=1, row=16)
    combobox17 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_17)
    combobox17['values'] = options
    combobox17.grid(column=1, row=17)
    combobox18 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_18)
    combobox18['values'] = options
    combobox18.grid(column=4, row=1)
    combobox19 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_19)
    combobox19['values'] = options
    combobox19.grid(column=4, row=2)
    combobox20 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_20)
    combobox20['values'] = options
    combobox20.grid(column=4, row=3)
    combobox21 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_21)
    combobox21['values'] = options
    combobox21.grid(column=4, row=4)
    combobox22 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_22)
    combobox22['values'] = options
    combobox22.grid(column=4, row=5)
    combobox23 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_23)
    combobox23['values'] = options
    combobox23.grid(column=4, row=6)
    combobox24 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_24)
    combobox24['values'] = options
    combobox24.grid(column=4, row=7)
    combobox25 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_25)
    combobox25['values'] = options
    combobox25.grid(column=4, row=8)
    combobox26 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_26)
    combobox26['values'] = options
    combobox26.grid(column=4, row=9)
    combobox27 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_27)
    combobox27['values'] = options
    combobox27.grid(column=4, row=10)
    combobox28 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_28)
    combobox28['values'] = options
    combobox28.grid(column=4, row=11)
    combobox29 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_29)
    combobox29['values'] = options
    combobox29.grid(column=4, row=12)
    combobox30 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_30)
    combobox30['values'] = options
    combobox30.grid(column=4, row=13)
    combobox31 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_31)
    combobox31['values'] = options
    combobox31.grid(column=4, row=14)
    combobox32 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_32)
    combobox32['values'] = options
    combobox32.grid(column=4, row=15)
    combobox33 = ttk.Combobox(master, width=25, textvariable=PRI_FIELD_33)
    combobox33['values'] = options
    combobox33.grid(column=4, row=16)
    lx0 = Label(master, text="\n\n", bg='#FDEBD0').grid(row=18, column=1)
    lx1 = Label(master, text="\n\n", bg='#FDEBD0').grid(row=20, column=1)
    lx2 = Label(master, text="\t\t", bg='#FDEBD0').grid(row=18, column=5)
    lx3 = Label(master, text="\t", bg='#FDEBD0').grid(row=18, column=2)
    button3 = Button(master, text="IMPORT", command=ImportCSV, state="normal", width=15).grid(row=8, column=6)
    button6 = Button(master, text="Back", command=back, state="normal", width=25).grid(row=21, column=1)
    button1 = Button(master, text="MAP", command=getmapping, width=10, state="normal").grid(row=19, column=1)
    if map_status==0:
        button2 = Button(master, text="SAVE", command=CreateCSV,state="disabled",width=10).grid(row=19, column=4)
        button4 = Button(master, text="CONTINUE", command=Conti, state="disabled", width=25).grid(row=21, column=4)
    master.mainloop()

#Needed because it worked this way
def back():
    master.destroy()
    homescreen()

def Rev_num():
    global rev_number
    rev_number = rev_number1.get()
    rev.destroy()

#ask for rev check , proceed to get IODB from user
def getinputs():
    home.destroy()
    global filename1, filename2
    global df_Cabinet_name, df_Alarm, df_PDP,df,df1
    global x
    global rev_number,rev_number1
    global rev

    error_check = 0
    x = mb.askyesno("Ques!", "Do you want to compare Revision?")
    if x == 1:
        file_get_2()
        try:
            df1 = pd.read_excel(filename1,'IODB')
            df = pd.read_excel(filename2,'IODB')
        except FileNotFoundError:
            mb.showerror("Error", "Please Select Files")
            homescreen()

        try:
            # CAbinet name sheet
            df_Cabinet_name = pd.read_excel(filename2, 'CabinetName')
            try:
                df_Cabinet_name['REV_DATE'] = df_Cabinet_name['REV_DATE'].dt.date
            except:
                mb.showwarning("Warning", 'CabinetName Sheet Does not contain Rev_Date Column!')
        except:
            mb.showerror("Error", 'CabinetName Sheet Not Found!')
            error_check = 1
        try:
            # alarm
            df_Alarm = pd.read_excel(filename2, 'Alarm')
        except:
            mb.showerror("Error", 'Alarm Sheet Not Found!')
            error_check = 1
        try:
            # pdp
            df_PDP = pd.read_excel(filename2, 'PDP')
        except:
            mb.showerror("Error", 'PDP Sheet Not Found!')
            error_check = 1
        if error_check==1:
            homescreen()
        else:
            compare()
            rev = Tk()
            rev.title('Revision Number')
            rev.geometry('300x300')
            rev_number1 = Entry(rev)
            temp=rev_number1.get()
            l = ttk.Label(rev, text="Revision Number").grid(column=0, row=0)
            rev_number1.grid(row=1, column=0)
            button1 = ttk.Button(rev, text="Rev. NO.", command=Rev_num, width=10, state="normal").grid(row=2,
                                                                                                            column=0)

            if headermatch == 1:
                mapping()
    else:
        file_get_1()
        try:
            df = pd.read_excel(filename1, 'IODB')
        except FileNotFoundError:
            mb.showerror("Error", "Please Select a File")
            exit()
        try:
            # CAbinet name sheet
            df_Cabinet_name = pd.read_excel(filename1, 'CabinetName')
            try:
                df_Cabinet_name['REV_DATE'] = df_Cabinet_name['REV_DATE'].dt.date
            except:
                mb.showwarning("Warning", 'CabinetName Sheet Does not contain Rev_Date Column!')

        except:
            mb.showerror("Error", 'CabinetName Sheet Not Found!')
            error_check = 1
        try:
            # alarm
            df_Alarm = pd.read_excel(filename1, 'Alarm')
        except:
            mb.showerror("Error", 'Alarm Sheet Not Found!')
            error_check = 1
        try:
            # pdp
            df_PDP = pd.read_excel(filename1, 'PDP')
        except:
            mb.showerror("Error", 'PDP Sheet Not Found!')
            error_check = 1
        if error_check == 1:
            homescreen()
        mapping()

#homescreen to get inputs
def homescreen():
    global test,enable
    global home
    test='bad'
    enable='bad'
    try:
        master.destroy()
    except:
        print("")
    try:
        exporttk.destroy()
    except:
        print("")
    try:
        crtpjt.destroy()
    except:
        print("")
    try:
        home.destroy()
    except:
        print("")


    home = Tk()
    home.title("HOME")
    home.geometry('500x600')
    home.configure(bg='#D6EAF8')
    l = Label(home, text="Lets Get Started",bg='#D6EAF8').grid(column=0, row=0, columnspan=2, pady=10,padx=50)
    l1 = Label(home, text="Project:", bg='#D6EAF8').grid(column=0, row=1, columnspan=2,pady=10)

    project =Button(home, text="Create Project", command=createpjt, width=30, bg='#FADBD8', state="normal").grid(row=2, column=1)
    select = Button(home, text="Select Project", command=selectepjt, width=30, bg='#FADBD8', state="normal").grid(row=2,
                                                                                                                  column=0)
    l2=Label(home, text="\nStart:", bg='#D6EAF8').grid(column=0, row=3, columnspan=2,pady=10)
    getfile = Button(home, text="Get Inputs", command=getinputs, width=30, bg='#AED6F1', state="disabled").grid(row=4,
                                                                                                              column=0,
                                                                                                              padx=20)
    close = Button(home, text="Exit", command=exit, width=30, bg='#FADBD8', state="normal").grid(row=4, column=1)
    name = Label(home, text="\n\n\n\nSMART DRAW IODB GEN TOOL", bg='#D6EAF8', font=("Courier", 15)).grid(column=0,
                                                                                                         row=5,
                                                                                                         columnspan=2,
                                                                                                         pady=50,
                                                                                                         padx=50)
    version = Label(home, text="V-0.01.05", bg='#D6EAF8', font=("Courier", 10)).grid(column=0, row=6, columnspan=2,padx=50)
    if enable=='ok':
        getfile = Button(home, text="Get Inputs", command=getinputs, width=30, bg='#FADBD8', state="normal").grid(
            row=4,
            column=0,
            padx=20)

    home.mainloop()

def selectepjt():
    global pjt_path,csv_path,excel_path
    pjt_path = filedialog.askdirectory(initialdir='C:\\')

    if pjt_path == '':
        mb.showerror("Error", "Please Select Valid Project Folder")
    else:
        csv_path = pjt_path + '\\Mapping'
        excel_path = pjt_path + '\\Export'
        os.makedirs(csv_path, exist_ok=True)
        os.makedirs(excel_path, exist_ok=True)
        mb.showinfo("Info","Project has been Selected.")

        getfile = Button(home, text="Get Inputs", command=getinputs, width=30, bg='#FADBD8', state="normal").grid(row=4,
                                                                                                                  column=0,
                                                                                                                  padx=20)
        home.update()


def createfld():
    global pjt_path,csv_path,excel_path,enable
    pjtname=Foldername.get()
    if folderpath_1!='' and pjtname!='':
        pjt_path = folderpath_1 + '\\' + pjtname
        print(pjt_path)
        os.makedirs(pjt_path,exist_ok=True)
        os.makedirs(pjt_path + '\\'+'Mapping', exist_ok=True)
        os.makedirs(pjt_path + '\\' + 'Export', exist_ok=True)
        csv_path = pjt_path + '\\Mapping'
        excel_path = pjt_path + '\\Export'
        crtpjt.destroy()
        mb.showinfo("Info","Your Project has been Created.")
        getfile = Button(home, text="Get Inputs", command=getinputs, width=30, bg='#FADBD8', state="normal").grid(row=4,
                                                                                                                  column=0,
                                                                                                                  padx=20)
        home.update()

    else:
        mb.showerror("Error","Please Select Valid Folder and Enter valid Name")
        crtpjt.destroy()
        createpjt()


def select_fld():
    global folderpath_1
    global Folderpath
    global crtpjt
    folderpath_1 = filedialog.askdirectory(initialdir='C:\\')
    Folderpath.insert(0,folderpath_1)
    crtpjt.update()


def createpjt():
    global Folderpath,Foldername,folderpath_1
    global test
    global crtpjt
    folderpath_1=''
    crtpjt =Tk()
    crtpjt.title('Create Project')
    crtpjt.geometry('450x250')
    crtpjt.configure(bg='#FBEEE6')
    Foldername = Entry(crtpjt)
    Foldername.insert(0,'')
    Foldername.grid(column=1, row=0,ipadx=35)
    Folderpath = Entry(crtpjt)
    Folderpath.insert(0, '')
    Folderpath.grid(column=1, row=1,ipadx=35,ipady=1)
    l = Label(crtpjt, text="Project Name",bg='#FBEEE6').grid(column=0, row=0,padx=20,pady=20)
    l3 = Label(crtpjt, text="  ",bg='#FBEEE6').grid(column=1, row=3)
    l1 = Label(crtpjt, text="\n\nProject Path\n\n",bg='#FBEEE6').grid(column=0, row=1)
    button1= Button(crtpjt, text="Back", command=homescreen, width=10, state="normal",bg='#AED6F1').grid(row=2, column=1)
    button2 = Button(crtpjt, text="Create Project", command=createfld, width=12, state="normal",bg='#AED6F1').grid(row=2, column=0)
    button3 = Button(crtpjt, text="Browse", command=select_fld, width=10, state="normal",bg='#AED6F1').grid(row=1, column=4,padx=10)
    crtpjt.mainloop()

# main Program Starts here:
if __name__ == '__main__':
    w1 = xlsxwriter.Workbook('hello.xlsx')
    homescreen()
