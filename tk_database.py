"""
/*!
 * pandas
 *
 * @Copyright (c) 2008-2011, AQR Capital Management, LLC, Lambda Foundry, Inc. and PyData Development Team All rights reserved.
 * @Copyright (c) 2011-2023, Open source contributors.
 *
 * https://pandas.pydata.org/pandas-docs/stable/getting_started/overview.html
 */
"""

"""
/*!
 * openpyxl
 *
 * © Copyright 2010 - 2024, Eric Gazoni, Charlie Clark
 *
 * https://openpyxl.readthedocs.io/en/stable/
 */
"""

from tkinter import *
from tkinter import ttk
from tkinterdnd2 import*
import tkinter.messagebox as tmsg
from tkinter import filedialog
import sqlite3
import os
import pandas as pd
import re
from datetime import datetime
import sys
import shutil
import zipfile

#バージョン管理番号
Version = "Ver(0.0.1)"

#主テーブル作成(装置idと装置名)
def create_main_table(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Machines(id TEXT PRIMARY KEY, name TEXT, customer TEXT, modelnum TEXT, enginetype TEXT, loginname TEXT, updatedate TEXT)
        ''')
        print("Table created successfully")
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブル(laiza)の作成(Headのバージョン)
def create_subtable_laiza(conn):
    try:
        conn.execute('''
                     CREATE TABLE IF NOT EXISTS Heads_laiza(machines_id INTEGER, headnum INTEGER, direction TEXT, unit TEXT, vpb TEXT, hpb TEXT, app TEXT, vpfpga TEXT, ipfpga TEXT, hpfpga TEXT,
                                                       ftp TEXT, xtp TEXT, amp TEXT, ma TEXT, dmdunit TEXT, dmdtype TEXT, dmdversion TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_laiza created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブル(livera)の作成(Headのバージョン)
def create_subtable_livera(conn):
    try:
        conn.execute('''
                     CREATE TABLE IF NOT EXISTS Heads_livera(machines_id INTEGER, headnum INTEGER, direction TEXT, unit TEXT, br TEXT, bsv TEXT, app TEXT, fpga TEXT, ftp TEXT, xtp TEXT,
                                                       amp TEXT, ma TEXT, dmdunit TEXT, dmdbr TEXT, dmdfpga TEXT, dmdtype TEXT, dmdversion TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_livera created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブル(ilia)の作成(Headのバージョン)
def create_subtable_ilia(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Heads_ilia(machines_id TEXT, headnum INTEGER, direction TEXT, unit TEXT, uibr TEXT, hpbr TEXT, app TEXT, isv TEXT, pfr TEXT, dmdunit TEXT,
                                                    dmdtype TEXT, dmdddc TEXT, tp TEXT, amtp TEXT, ma TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_ilia created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
        
#副テーブル(lacs)の作成(Headのバージョン)
def create_subtable_lacs(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Heads_lacs(machines_id TEXT, headnum INTEGER, direction TEXT, unit TEXT, unirev TEXT, bootsw TEXT, softv TEXT, fpgarev TEXT, ddhard TEXT,
                                                          xtp TEXT, amp TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_lacs created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
        
#副テーブル(PE)の作成(Headのバージョン)
def create_subtable_pe(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Heads_pe(machines_id TEXT, headnum INTEGER, direction TEXT, bootsw TEXT, cpufpga TEXT, mccon TEXT, revs TEXT, mcdoft TEXT, hardrev TEXT, fpgahrev TEXT, tpr TEXT,
                                                          loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_pe created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
        
#副テーブル(MATE)の作成(Headのバージョン)
def create_subtable_mate(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Heads_mate(machines_id TEXT, headnum INTEGER, direction TEXT, ut TEXT, fpgabr TEXT, dmdbr TEXT, app TEXT, isoc TEXT, pfpga TEXT, hfpga TEXT, dmdunit TEXT,
                                                           dmdtype TEXT, dlpcver TEXT, tp TEXT, amp TEXT, mc TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_mate created successfully")
    except sqlite3.Error as e:
        print(f"An error occurred: {e}") 
        
#副テーブル(dpc)の作成(DPCのバージョン)
def create_subtable_dpc(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS DPC(machines_id TEXT, dataconver TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_DPC created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブル(dxpcon)の作成(DXPCONのバージョン)
def create_subtable_dxpcon(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS DXPCON(machines_id TEXT, dxpconver TEXT, loginname TEXT, updatedate TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_DXPCON created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
             
#副テーブル(Excelファイル読み込み結果)の作成(Basic Infomation)
def create_subtable_basicinfo(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS BasicInfo(machines_id TEXT, machines_type TEXT, customer TEXT, modelnum TEXT, enginetype TEXT)
        ''')
        print("SUb_Table_BasicInfo created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブル(Excelファイルの読み込みリスト)の作成(Excel List)
def create_subtable_excellist(coon):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Excellist(excelname TEXT, regtime TEXT, latestflag TEXT)
        ''')
        print("SUb_Table_Excellist created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
        
#副テーブルの作成(ユーザー)
def create_subtable_user(conn):
    try:
        conn.execute('''
                    CREATE TABLE IF NOT EXISTS Users(id TEXT, passward TEXT, name TEXT)
        ''')
        print("SUb_Table_Users created successfully")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")
        
#主テーブルに登録        
def insert_machine_id_if_not_exists(conn, machine_id, machine_name, customer, modelnum, enginetype, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # プライマリキーの存在チェック
        cur.execute("SELECT 1 FROM Machines WHERE id = ?", (machine_id,))
        if cur.fetchone() is None:
            # プライマリキーが存在しない場合、レコードを追加
            cur.execute("INSERT INTO Machines (id, name, customer, modelnum, enginetype, loginname, updatedate) VALUES (?, ?, ?, ?, ?, ?, ?)", (machine_id, machine_name, customer, modelnum, enginetype, loginname, updatedate))
            conn.commit()
            print("Machine added successfully")
        else:
            print("Machine ID already exists. No record added.")
    except sqlite3.Error as e:
        tmsg.showerror("Error",f"An error occurred: {e}")

#副テーブルに登録(laiza)
def insert_heads_laiza(conn, machine_id, headnum, direction, unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # Head情報の存在チェック
        cur.execute('''SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND direction = ?''', (machine_id, headnum, direction))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_laiza (machines_id, headnum, direction, unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                ( machine_id, headnum, direction, unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_laiza data added successfully")
            msg = "Heads_laiza data added successfully"
            return True,msg
        else:
            print("Heads_laiza already registered")
            msg = "Heads_laiza already registered"
            return False,msg
    except sqlite3.Error as e:
        msg = f"An error occurred: {e}"
        return False,msg
    
#副テーブルに登録(livera)
def insert_heads_livera(conn, machine_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # Head情報の存在チェック
        cur.execute('''SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND direction = ?''', (machine_id, headnum, direction))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_livera (machines_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        ( machine_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_livera data added successfully")
            msg = "Heads_livera data added successfully"
            return True,msg
        else:
            print("Heads_livera already registered")
            msg = "Heads_livera already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブルに登録(ilia)
def insert_heads_ilia(conn, machine_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # Head情報の存在チェック
        cur.execute('''SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ? AND direction = ?''', (machine_id, headnum, direction))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_ilia (machines_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        ( machine_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_ilia data added successfully")
            msg = "Heads_ilia data added successfully"
            return True,msg
        else:
            print("Heads_ilia already registered")
            msg = "Heads_ilia already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブルに登録(lacs)
def insert_heads_lacs(conn, machine_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # Head情報の存在チェック
        cur.execute('''SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND direction = ?''', (machine_id, headnum, direction))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_lacs (machines_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (machine_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_lacs data added successfully")
            msg = "Heads_lacs data added successfully"
            return True,msg
        else:
            print("Heads_lacs already registered")
            msg = "Heads_lacs already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブルに登録(PE)
def insert_heads_pe(conn, machine_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # Head情報の存在チェック
        cur.execute('''SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? ''', (machine_id, headnum))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_pe (machines_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (machine_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_pe data added successfully")
            msg = "Heads_pe data added successfully"
            return True,msg
        else:
            msg = "Heads_pe already registered"
            return False,msg
    except sqlite3.Error as e:
         print(f"An error occurred: {e}")   

#副テーブルに登録(mate)
def insert_heads_mate(conn, machine_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #Head情報の損じチェック
        cur.execute('''SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? ''', (machine_id, headnum))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Heads_mate (machines_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (machine_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, loginname, updatedate, "Latest"))
            conn.commit()
            print("Heads_mate data added successfully")
            msg = "Heads_mate data added successfully"
            return True,msg
        else:
            msg = "Heads_pe already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")     

#副テーブルに登録(DPC)
def insert_dpc(conn, machine_id, dpcver, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #DPCバージョンの存在チェック
        cur.execute('''SELECT * FROM DPC WHERE machines_id = ?''',(machine_id,))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO DPC (machines_id, dataconver, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?)",(machine_id, dpcver, loginname, updatedate, "Latest"))
            conn.commit()
            print("datacon_version data added successfully")
            msg = "datacon_version data added successfully"
            return True,msg
        else:
            print("datacon_version data already registered")
            msg = "datacon_version data already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブルに登録(DXPCON)
def insert_dxpcon(conn, machine_id, dxpconver, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #DXPCONバージョンの存在チェック
        cur.execute('''SELECT * FROM DXPCON WHERE machines_id = ?''',(machine_id,))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO DXPCON (machines_id, dxpconver, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?)",(machine_id, dxpconver, loginname, updatedate, "Latest"))
            conn.commit()
            print("dxpcon_version data added successfully")
            msg = "dxpcon_version data added successfully"
            return True,msg
        else:
            print("dxpcon_version data already registered")
            msg = "dxpcon_version data already registered"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
 
 #副テーブルに登録(BasicInfo)
def insert_basicinfo(conn, machines_id, machines_type, customer, modelnum, enginetype):
    try:
        cur=conn.cursor()
        #装置IDの存在チェック
        cur.execute('''SELECT * FROM BasicInfo WHERE machines_id = ?''',(machines_id,))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO BasicInfo (machines_id, machines_type, customer, modelnum, enginetype) VALUES (?, ?, ?, ?, ?)",(machines_id, machines_type, customer, modelnum, enginetype))
            conn.commit()
        else:
            return
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")       

#副テーブルの登録・更新(Excellist)
def insert_excellist(conn, excelname, timestamp):
    try:
        cur=conn.cursor()
        cur.execute('''SELECT * FROM Excellist WHERE latestflag = ?''',("Latest",))
        info = cur.fetchone()
        if info:
            cur.execute('''UPDATE Excellist SET latestflag = 'Older' WHERE latestflag = ?''', ("Latest",))
            cur.execute("INSERT INTO Excellist (excelname, regtime, latestflag) VALUES (?, ?, ?)",(excelname,timestamp, "Latest"))
            conn.commit()
        else:
            cur.execute("INSERT INTO Excellist (excelname, regtime, latestflag) VALUES (?, ?, ?)",(excelname,timestamp, "Latest"))
            conn.commit()
        print("Excellist added successfully")
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブルの登録(Users)
def insert_users(conn,id,username):
    try:
        cur=conn.cursor()
        cur.execute('''SELECT * FROM Users WHERE id = ?''',(id,))
        if cur.fetchone() is None:
            cur.execute("INSERT INTO Users (id, passward, name) VALUES(?, ?, ?)",(id,"hirai285",username))
            conn.commit()
            print("User recogniazed suceess")
        else:
           return
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#副テーブル削除(BasicInfo)
def delete_basicinfo(conn):
    try:
        cur=conn.cursor()
        cur.execute('''DROP TABLE BasicInfo''')
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        
#エクセルファイルの更新判定
def inquiry(conn,timestamp):
    cur=conn.cursor()
    cur.execute("SELECT * FROM  Excellist WHERE regtime = ? AND latestflag = ?", (timestamp,"Latest"))
    info = cur.fetchone()
    if info is None:
        return False
    else :
        return True
         
#IDによるエンジン情報表示
def display_machine_info(conn,machine_id):
    # Machinesテーブルから装置情報を取得
    if machine_id != '' and searchsettei.get() == 0:
        print("OK")
    elif searchsettei.get() == 1:
        print("OK")
    elif searchsettei.get() == 2:
        print("OK")
    else:
        tmsg.showerror("Error","製番が空白です")
        return
    cur=conn.cursor()
    if(searchsettei.get() == 0 or searchsettei.get() == 1):
        if(searchsettei.get() == 0):
            cur.execute('SELECT * FROM  Machines WHERE id = ? ', (machine_id,))
        else:
            reroad_customerlist(cur)
            customer ,enginetype= show_customerandenginetype_dialog(root)
            if enginetype == "":
                tmsg.showerror("Error","ダイアログ入力エラー。エンジンを選択してOKをクリックしてください")
                return
            cur.execute(query_customerinfo,(('%' + customer + '%'), enginetype))
        info= cur.fetchall()
        if info:
            print("OK")
        else:
            if searchsettei.get() == 0:
                print(f"ProductID {machine_id} does not exist.")
                tmsg.showerror("Error",f"製番 {machine_id} が存在しません")
                return
            else:
                print("Can't find any information.")
                tmsg.showerror("Error","情報が見つかりません")
                return
        if info[0][4] == "LAIZA":
            global tree_common
            global tree_separate
            global scrollbarx
            global scrollbary
            global sort_customer
            sort_customer = True
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None   
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None 
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=laiza_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no') 
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no') 
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('LI Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Vector Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Head Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LI App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Vector Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Intersection Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Head Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI FHD Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI XGA Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Area Mask Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LI Unit Type',anchor='center', text='LI Unit Type')
            tree_separate.heading('LI Vector Process Board Revision', anchor='center', text='LI Vector Process Board Revision')
            tree_separate.heading('LI Head Process Board Revision', anchor='center', text='LI Head Process Board Revision')
            tree_separate.heading('LI App Software Version', anchor='center', text='LI App Software Version')
            tree_separate.heading('LI Vector Process FPGA Version', anchor='center', text='LI Vector Process FPGA Version')
            tree_separate.heading('LI Intersection Process FPGA Version', anchor='center', text='LI Intersection Process FPGA Version')
            tree_separate.heading('LI Head Process FPGA Version', anchor='center', text='LI Head Process FPGA Version')
            tree_separate.heading('LI FHD Test Pattern Version', anchor='center', text='LI FHD Test Pattern Version')
            tree_separate.heading('LI XGA Test Pattern Version', anchor='center', text='LI XGA Test Pattern Version')
            tree_separate.heading('LI Area Mask Pattern Version', anchor='center', text='LI Area Mask Pattern Version')
            tree_separate.heading('LI MAC Address', anchor='center', text='LI MAC Address')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
        elif info[0][4] == "LIVERA":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=livera_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('LV Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Boot Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV FHD Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV XGA Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Area Mask Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LV Unit Type', anchor='center', text='LV Unit Type')
            tree_separate.heading('LV Board Revision', anchor='center', text='LV Board Revision')
            tree_separate.heading('LV Boot Software Version', anchor='center', text='LV Boot Software Version')
            tree_separate.heading('LV App Software Version', anchor='center', text='LV App Software Version')
            tree_separate.heading('LV FPGA Version', anchor='center', text='LV FPGA Version')
            tree_separate.heading('LV FHD Test Pattern Version', anchor='center',text='LV FHD Test Pattern Version')
            tree_separate.heading('LV XGA Test Pattern Version', anchor='center', text='LV XGA Test Pattern Version')
            tree_separate.heading('LV Area Mask Pattern Version', anchor='center', text='LV Area Mask Pattern Version')
            tree_separate.heading('LV MAC Address', anchor='center', text='LV MAC Address')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Board Revision', anchor='center', text='DMD Board Revisions')
            tree_separate.heading('DMD FPGA Version', anchor='center', text='DMD FPGA Version')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
        elif info[0][4] == "ILIA":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定   
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定                 
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=ilia_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0' ,width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('IA Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('IA USB Interface Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Head Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Interface SoC Version', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Plot FPGA Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('IA Unit Type', anchor='center', text='IA Unit Type')
            tree_separate.heading('IA USB Interface Board Revision', anchor='center', text='IA USB Interface Board Revision')
            tree_separate.heading('IA Head Process Board Revision', anchor='center', text='IA Head Process Board Revision')
            tree_separate.heading('IA App Software Version', anchor='center', text='IA App Software Version')
            tree_separate.heading('IA Interface SoC Version', anchor='center', text='IA Interface SoC Version')
            tree_separate.heading('IA Plot FPGA Revision', anchor='center', text='IA Plot FPGA Revision')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100', anchor='center', text='DMD DDC4100')
            tree_separate.heading('IA Test Pattern Revision', anchor='center', text='IA Test Pattern Revision')
            tree_separate.heading('IA Area Mask Pattern Revision', anchor='center', text='IA Area Mask Pattern Revision')
            tree_separate.heading('IA MAC Address', anchor='center', text='IA MAC Address')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
        elif info[0][4] == "LACS":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定   
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定                 
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=lacs_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('LE_Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Unit Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Boot Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_FPGA Hardware Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_DD Hardware Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_XGA Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0',text='')
            tree_common.heading('ID', anchor='center', text='Product ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0',text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LE_Unit Type',anchor='center', text='LE_Unit Type')
            tree_separate.heading('LE_Unit Revision', anchor='center', text='LE_Unit Revision')
            tree_separate.heading('LE_Boot Software Version', anchor='center', text='LE_Boot Software Version')
            tree_separate.heading('LE_Software Version', anchor='center', text='LE_Software Version')
            tree_separate.heading('LE_FPGA Hardware Revision', anchor='center', text='LE_FPGA Hardware Revision')
            tree_separate.heading('LE_DD Hardware Revision', anchor='center', text='LE_DD Hardware Revision')
            tree_separate.heading('LE_XGA Test Pattern Revision', anchor='center', text='LE_XGA Test Pattern Revision')
            tree_separate.heading('LE_Area Mask Pattern Revision', anchor='center', text='LE_Area Mask Pattern Revision')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
        elif info[0][4] == "PE" or info[0][4] == "PE-Ver2":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定   
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定                 
            tree_common = ttk.Treeview(mainframe1,columns=common_column)
            tree_separate = ttk.Treeview(mainframe1,columns=pe_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('PE_Boot SoftVersion',anchor='center', width=200, stretch='no')
            tree_separate.column('PE-CPUCoreFPGA_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
            tree_separate.column('IE_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
            tree_separate.column('RE-VS_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE-FPGA_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE_test pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0',text='')
            tree_common.heading('ID', anchor='center', text='Product ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0',text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('PE_Boot SoftVersion',anchor='center', text='PE_Boot SoftVersion')
            tree_separate.heading('PE-CPUCoreFPGA_hard Revision', anchor='center', text='PE-CPUCoreFPGA_hard Revision')
            tree_separate.heading('IE_MC Control SoftVersion', anchor='center', text='IE_MC Control SoftVersion')
            tree_separate.heading('IE_hard Revision', anchor='center', text='IE_hard Revision')
            tree_separate.heading('PE_MC Control SoftVersion', anchor='center', text='PE_MC Control SoftVersion')
            tree_separate.heading('RE-VS_hard Revision', anchor='center', text='RE-VS_hard Revision')
            tree_separate.heading('PE-FPGA_hard Revision', anchor='center', text='PE-FPGA_hard Revision')
            tree_separate.heading('PE_test pattern Revision', anchor='center', text='PE_test pattern Revision')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
        elif info[0][4] == 'MATE' or info[0][4] == "MATE3":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定   
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定                 
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=mate_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('MATE Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('MATE FPGA Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE DMD Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Interface SoC Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Plot FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Head FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DLPC910 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0',text='')
            tree_common.heading('ID', anchor='center', text='Product ID')
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0',text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('MATE Unit Type',anchor='center', text='MATE Unit Type')
            tree_separate.heading('MATE FPGA Board Revision', anchor='center', text='MATE FPGA Board Revision')
            tree_separate.heading('MATE DMD Board Revision', anchor='center', text='MATE DMD Board Revision')
            tree_separate.heading('MATE App Software Version', anchor='center', text='MATE App Software Version')
            tree_separate.heading('MATE Interface SoC Version', anchor='center', text='MATE Interface SoC Version')
            tree_separate.heading('MATE Plot FPGA Version', anchor='center', text='MATE Plot FPGA Version')
            tree_separate.heading('MATE Head FPGA Version', anchor='center', text='MATE Head FPGA Version')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DLPC910 Version', anchor='center', text='DMD DLPC910 Version')
            tree_separate.heading('MATE Test Pattern Revision', anchor='center', text='MATE Test Pattern Revision')
            tree_separate.heading('MATE Area Mask Pattern Revision', anchor='center', text='MATE Area Mask Pattern Revision')
            tree_separate.heading('MATE MAC Address', anchor='center', text='MATE MAC Address')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='') 
        count =0
        for loop in info:       
            cur.execute('SELECT * FROM DPC Where machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
            dpcver = cur.fetchone()
            if dpcver:
                print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
            else:
                dpcver = ['----' for temp in range(2)]
            cur.execute('SELECT * FROM DXPCON Where machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
            dxpconver = cur.fetchone()
            if dxpconver:
                print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
            else:
                dxpconver = ['----' for temp in range(2)]
            # Headsテーブルから装置IDに関連する部品情報を取得
            if loop[4] == "LAIZA":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
            elif loop[4] == "LIVERA":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0] ,0, "Latest"))
            elif loop[4] == "ILIA":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ?  AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
            elif loop[4] == "LACS":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
            elif loop[4] == "PE" or loop[4] == "PE-Ver2":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
            elif loop[4] == "MATE" or loop[4] == "MATE3":
                if settei.get() == 0 or searchsettei.get() == 0:
                    cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
            headsinfo = cur.fetchall()
            if headsinfo:
                print("Heads Information exsit:")
            else:
                headsinfo = [['----']*17]
            if loop[4] == "LAIZA":
                #レコードの追加
                for head in headsinfo:
                    cur.execute(query_laiza,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----'] 
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], updatedate[0], updatedate[1], loop[2], ""))
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
            elif loop[4] == "LIVERA":
                #レコードの追加
                for head in headsinfo:
                    cur.execute(query_livera,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----'] 
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], updatedate[0], updatedate[1], loop[2], ""))  
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
            elif loop[4] == "ILIA":
                #レコードの追加
                for head in headsinfo: 
                    cur.execute(query_ilia,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----']
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], updatedate[0], updatedate[1], loop[2], ""))
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
            elif loop[4]== "LACS":
                #レコードの追加
                for head in headsinfo: 
                    cur.execute(query_lacs,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----']
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], updatedate[0], updatedate[1], loop[2], ""))
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
            elif loop[4] == "PE" or loop[4] == "PE-Ver2":
                #レコードの追加
                for head in headsinfo: 
                    cur.execute(query_pe,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----']
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], updatedate[0], updatedate[1], loop[2], ""))
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
            elif loop[4] == "MATE" or loop[4] == "MATE3":
                #レコードの追加
                for head in headsinfo: 
                    cur.execute(query_mate,(head[1], loop[0], loop[0], loop[0]))
                    updatedate = cur.fetchone()
                    if updatedate:
                        print(f"Updatedate - Date: {updatedate[0]}")  
                    else:
                        updatedate =  ['----']
                    tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                    tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], updatedate[0], updatedate[1], loop[2], ""))
                    count +=1
                # Styleの設定
                style = ttk.Style()
                # Treeviewの選択時の背景色をデフォルトと同じにする
                style.map('Treeview', 
                        background=[('selected', style.lookup('Treeview', 'background'))],
                        foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                style.configure("Treeview.Heading", font=(None, 12))
                # Treeviewの枠線を非表示にする
                style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                #横スクロールバーの追加
                scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
                tree_separate.configure(xscrollcommand = scrollbarx.set)
                scrollbarx[ 'command' ] = tree_separate.xview
                scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                # 縦スクロールバー
                scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
                tree_separate.configure(yscroll = scrollbary.set)
                scrollbary[ 'command' ] = sync_tree_separate_yview
                scrollbary.place(relx=0.985, relheight=0.98)
                #マウスホイールの同期
                tree_common.bind("<MouseWheel>", on_mouse_wheel)
                tree_separate.bind("<MouseWheel>", on_mouse_wheel)
                # ウィジェットの配置
                tree_common.place(relheight=1.0,relwidth=0.52)
                tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
                save_button['state'] = 'normal'
                reload()
    else:
        modelnum = show_modelnumonly_dialog(root)
        if scrollbarx != None:
            scrollbarx.pack_forget()
            scrollbarx = None    
        if scrollbary != None:
            scrollbary.pack_forget()
            scrollbary = None
        if tree_common != None:
            tree_common.destroy()
            tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
        if tree_separate != None:
            tree_separate.destroy()
            tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定
        tree_common = ttk.Treeview(mainframe1,columns=main_column)
        tree_common.bind("<<TreeviewSelect>>", select_record)
        #列の設定
        tree_common.column('#0',width=0, stretch='no')
        tree_common.column('ID', anchor='center', width=200, stretch='no')
        tree_common.column('Model', anchor='center', width=200, stretch='no')
        tree_common.column('Customer', anchor='center', width=250, stretch='no')
        tree_common.column('Machine_No', anchor='center', width=200, stretch='no')
        tree_common.column('Engine', anchor='center', width=200, stretch='no')
        tree_common.column('blank', anchor='center', width=100, stretch='no')   
        #列の見出し
        tree_common.heading('#0',text='')
        tree_common.heading('ID', anchor='center', text='ID',command=lambda _col='ID': \
                        treeview_main_sort_column(tree_common, _col, False))
        tree_common.heading('Model', anchor='center', text='Model',command=lambda _col='Model': \
                        treeview_main_sort_column(tree_common, _col, False))
        tree_common.heading('Customer', anchor='center', text='Customer',command=lambda _col='Customer': \
                        treeview_main_sort_column(tree_common, _col, False))
        tree_common.heading('Machine_No', anchor='center', text='Machine_No',command=lambda _col='Machine_No': \
                        treeview_main_sort_column(tree_common, _col, False))
        tree_common.heading('Engine', anchor='center', text='Engine',command=lambda _col='Engine': \
                        treeview_main_sort_column(tree_common, _col, False))
        tree_common.heading('blank', anchor='center',text='')  
        cur.execute('SELECT * FROM  Machines Where  modelnum LIKE ? ORDER BY id ASC',( modelnum ,))
        result = cur.fetchall()
        count = 0
        tree_common.tag_configure("red", foreground='red')
        if result:
            for info in result:
                tree_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], info[2], info[3], info[4], ""))
                count +=1
        else:
            print("infomation nothing")
        cur.execute('''SELECT * From BasicInfo WHERE machines_id NOT IN (SELECT id From Machines) and modelnum LIKE ? ORDER BY machines_id ASC''',(modelnum,))
        subresult = cur.fetchall()
        if subresult:
            for subinfo in subresult:
                tree_common.insert(parent='', index='end', iid=count, values=(subinfo[0], subinfo[1], subinfo[2], subinfo[3], subinfo[4], ""), tags="red")
                count +=1
        else:
            print("subinfomation nothing")
        
        # Styleの設定
        style_id = ttk.Style()
        style_id.map('Treeview',
                    foreground=[('selected', 'white')],
                    background=[('selected', 'deepskyblue')])
        style_id.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
        style_id.configure("Treeview.Heading", font=(None, 12))
        # Treeviewの枠線を非表示にする
        style_id.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
        # 縦スクロールバー
        scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL, command=tree_common.yview)
        tree_common.configure(yscroll = scrollbary.set)
        scrollbary.place(relx=0.986, relheight=0.98)
        #横スクロールバー
        scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL, command=tree_common.xview)
        tree_common.configure(xscroll = scrollbarx.set)
        scrollbarx.place(rely=0.98, relwidth=0.986)
        save_button['state'] = 'normal'
        # ウィジェットの配置
        tree_common.place(relheight=1.0,relwidth=1.0)

#-----------treeview　Yスクロールバー同期処理---------#
def sync_tree_separate_yview(*args):
    tree_separate.yview(*args)
    # tree2のスクロールイベントに応じてtree1のスクロールを同期
    tree_common.yview_moveto(tree_separate.yview()[0]) 
    
def sync_tree_log_separete_yview(*args):
    tree_log_separete.yview(*args)
    # tree2のスクロールイベントに応じてtree1のスクロールを同期
    tree_log_common.yview_moveto(tree_log_separete.yview()[0]) 

#マウスホイールの禁止
def on_mouse_wheel(event):
    return "break"

#一覧の表示(装置種別)
def list_view(conn):
    cur=conn.cursor()
    machine_name = show_machinetype_selection_dialog(root)
    if machine_name == "":
        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
        return
    engine = ""
    if machine_name in laizalist:
        engine = "LAIZA"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in liveralist:
        engine = "LIVERA"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in ilialist:
        engine = "ILIA"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in lacslist:
        engine = "LACS"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in pe2list:
        engine = "PE-Ver2"
        if engine in machine_name:
            machine_name = separate_name(machine_name)
    elif machine_name in pelist:
        engine = "PE"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in matelist:
        engine = "MATE"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
    elif machine_name in mate3list:
        engine = "MATE3"
        if engine in machine_name.upper():
            machine_name = separate_name(machine_name)
            
    cur.execute('SELECT * FROM  Machines WHERE name = ? COLLATE NOCASE AND enginetype = ? ORDER BY id ASC', (machine_name, engine))
    info= cur.fetchall()
    if info:
        global tree_common
        global tree_separate
        global scrollbarx
        global scrollbary
        global sort_id
        global sort_customer
        sort_id = False
        sort_customer = False
        if engine == "LAIZA":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定   
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_separate = ttk.Treeview(mainframe1, columns=laiza_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no')
            tree_separate.column('LI Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('LI Vector Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Head Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LI App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Vector Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Intersection Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Head Process FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI FHD Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI XGA Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI Area Mask Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LI MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID', command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LI Unit Type',anchor='center', text='LI Unit Type')
            tree_separate.heading('LI Vector Process Board Revision', anchor='center', text='LI Vector Process Board Revision')
            tree_separate.heading('LI Head Process Board Revision', anchor='center', text='LI Head Process Board Revision')
            tree_separate.heading('LI App Software Version', anchor='center', text='LI App Software Version')
            tree_separate.heading('LI Vector Process FPGA Version', anchor='center', text='LI Vector Process FPGA Version')
            tree_separate.heading('LI Intersection Process FPGA Version', anchor='center', text='LI Intersection Process FPGA Version')
            tree_separate.heading('LI Head Process FPGA Version', anchor='center', text='LI Head Process FPGA Version')
            tree_separate.heading('LI FHD Test Pattern Version', anchor='center', text='LI FHD Test Pattern Version')
            tree_separate.heading('LI XGA Test Pattern Version', anchor='center', text='LI XGA Test Pattern Version')
            tree_separate.heading('LI Area Mask Pattern Version', anchor='center',text='LI Area Mask Pattern Version')
            tree_separate.heading('LI MAC Address', anchor='center', text='LI MAC Address')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:          
                    cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*17]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)]
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)]
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_laiza,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], updatedate[0], updatedate[1], loop[2], ""))                        
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #横スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            reload()
        elif engine == "LIVERA":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定               
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            tree_separate = ttk.Treeview(mainframe1, columns=livera_column)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no')
            tree_separate.column('LV Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Boot Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV FHD Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV XGA Test Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV Area Mask Pattern Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LV MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID',command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LV Unit Type',anchor='center', text='LV Unit Type')
            tree_separate.heading('LV Board Revision', anchor='center', text='LV Board Revision')
            tree_separate.heading('LV Boot Software Version', anchor='center', text='LV Boot Software Version')
            tree_separate.heading('LV App Software Version', anchor='center', text='LV App Software Version')
            tree_separate.heading('LV FPGA Version', anchor='center', text='LV FPGA Version')
            tree_separate.heading('LV FHD Test Pattern Version', anchor='center',text='LV FHD Test Pattern Version')
            tree_separate.heading('LV XGA Test Pattern Version', anchor='center', text='LV XGA Test Pattern Version')
            tree_separate.heading('LV Area Mask Pattern Version', anchor='center', text='LV Area Mask Pattern Version')
            tree_separate.heading('LV MAC Address', anchor='center', text='LV MAC Address')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Board Revision', anchor='center', text='DMD Board Revisions')
            tree_separate.heading('DMD FPGA Version', anchor='center', text='DMD FPGA Version')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:
                    cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*17]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)] 
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)] 
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_livera,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], updatedate[0], updatedate[1], loop[2], ""))
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            # reload()
        elif engine == "ILIA":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定               
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            tree_separate = ttk.Treeview(mainframe1, columns=ilia_column)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0' ,width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no')
            tree_separate.column('IA Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('IA USB Interface Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Head Process Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Interface SoC Version', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Plot FPGA Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DDC4100', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IA MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0',text='')
            tree_common.heading('ID', anchor='center', text='Product ID',command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0',text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('IA Unit Type',anchor='center', text='IA Unit Type')
            tree_separate.heading('IA USB Interface Board Revision', anchor='center', text='IA USB Interface Board Revision')
            tree_separate.heading('IA Head Process Board Revision', anchor='center', text='IA Head Process Board Revision')
            tree_separate.heading('IA App Software Version', anchor='center', text='IA App Software Version')
            tree_separate.heading('IA Interface SoC Version', anchor='center', text='IA Interface SoC Version')
            tree_separate.heading('IA Plot FPGA Revision', anchor='center', text='IA Plot FPGA Revision')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DDC4100', anchor='center', text='DMD DDC4100')
            tree_separate.heading('IA Test Pattern Revision', anchor='center',text='IA Test Pattern Revision')
            tree_separate.heading('IA Area Mask Pattern Revision', anchor='center', text='IA Area Mask Pattern Revision')
            tree_separate.heading('IA MAC Address', anchor='center', text='IA MAC Address')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:
                    cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*16]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)] 
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)] 
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_ilia,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], updatedate[0], updatedate[1], loop[2], ""))                     
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            reload()
        elif engine == "LACS":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定               
            tree_common = ttk.Treeview(mainframe1,columns=common_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            tree_separate = ttk.Treeview(mainframe1,columns=lacs_column)
            #列の設定
            tree_common.column('#0',width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0',width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no')
            tree_separate.column('LE_Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Unit Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Boot Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_FPGA Hardware Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_DD Hardware Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_XGA Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('LE_Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID',command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('LE_Unit Type',anchor='center', text='LE_Unit Type')
            tree_separate.heading('LE_Unit Revision', anchor='center', text='LE_Unit Revision')
            tree_separate.heading('LE_Boot Software Version', anchor='center', text='LE_Boot Software Version')
            tree_separate.heading('LE_Software Version', anchor='center', text='LE_Software Version')
            tree_separate.heading('LE_FPGA Hardware Revision', anchor='center', text='LE_FPGA Hardware Revision')
            tree_separate.heading('LE_DD Hardware Revision', anchor='center', text='LE_DD Hardware Revision')
            tree_separate.heading('LE_XGA Test Pattern Revision', anchor='center', text='LE_XGA Test Pattern Revision')
            tree_separate.heading('LE_Area Mask Pattern Revision', anchor='center', text='LE_Area Mask Pattern Revision')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:
                    cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*16]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)] 
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)] 
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_lacs,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], updatedate[0], updatedate[1], loop[2], ""))                     
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            reload()
        elif engine == "PE" or engine == "PE-Ver2":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定               
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            tree_separate = ttk.Treeview(mainframe1, columns=pe_column)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('PE_Boot SoftVersion',anchor='center', width=200, stretch='no')
            tree_separate.column('PE-CPUCoreFPGA_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('IE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
            tree_separate.column('IE_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
            tree_separate.column('RE-VS_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE-FPGA_hard Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('PE_test pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0', text='')
            tree_common.heading('ID', anchor='center', text='Product ID',command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0', text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('PE_Boot SoftVersion',anchor='center', text='PE_Boot SoftVersion')
            tree_separate.heading('PE-CPUCoreFPGA_hard Revision', anchor='center', text='PE-CPUCoreFPGA_hard Revision')
            tree_separate.heading('IE_MC Control SoftVersion', anchor='center', text='IE_MC Control SoftVersion')
            tree_separate.heading('IE_hard Revision', anchor='center', text='IE_hard Revision')
            tree_separate.heading('PE_MC Control SoftVersion', anchor='center', text='PE_MC Control SoftVersion')
            tree_separate.heading('RE-VS_hard Revision', anchor='center', text='RE-VS_hard Revision')
            tree_separate.heading('PE-FPGA_hard Revision', anchor='center', text='PE-FPGA_hard Revision')
            tree_separate.heading('PE_test pattern Revision', anchor='center', text='PE_test pattern Revision')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:
                    cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*16]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)] 
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)] 
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_pe,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], updatedate[0], updatedate[1], loop[2], ""))                     
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            reload()
        elif engine == "MATE" or engine == "MATE3":
            if scrollbarx != None:
                scrollbarx.pack_forget()
                scrollbarx = None    
            if scrollbary != None:
                scrollbary.pack_forget()
                scrollbary = None
            if tree_common != None:
                tree_common.destroy()
                tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
            if tree_separate != None:
                tree_separate.destroy()
                tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定               
            tree_common = ttk.Treeview(mainframe1, columns=common_column)
            tree_common.bind("<<TreeviewSelect>>", select_record)
            tree_separate = ttk.Treeview(mainframe1, columns=mate_column)
            #列の設定
            tree_common.column('#0', width=0, stretch='no')
            tree_common.column('ID', anchor='center', width=150, stretch='no')
            tree_common.column('Model', anchor='center', width=150, stretch='no')
            tree_common.column('HeadNum', anchor='center', width=100, stretch='no')
            tree_common.column('Direction', anchor='center', width=100, stretch='no')
            tree_separate.column('#0', width=0, stretch='no')
            tree_separate.column('Machine_No', anchor='center', width=200, stretch='no')
            tree_separate.column('DXPCON Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DPC Version', anchor='center', width=220, stretch='no') 
            tree_separate.column('MATE Unit Type',anchor='center', width=200, stretch='no')
            tree_separate.column('MATE FPGA Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE DMD Board Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE App Software Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Interface SoC Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Plot FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Head FPGA Version', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Unit Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD Type', anchor='center', width=200, stretch='no')
            tree_separate.column('DMD DLPC910 Version', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Test Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
            tree_separate.column('MATE MAC Address', anchor='center', width=200, stretch='no')
            tree_separate.column('Updater', anchor='center', width=200, stretch='no')
            tree_separate.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_separate.column('Customer', anchor='center', width=200, stretch='no')
            tree_separate.column('blank', anchor='center', width=100, stretch='no')
            #列の見出し
            tree_common.heading('#0',text='')
            tree_common.heading('ID', anchor='center', text='Product ID',command=lambda: treeview_sort_idcolumn(tree_common, tree_separate, 'ID'))
            tree_common.heading('Model', anchor='center', text='Model')
            tree_common.heading('HeadNum', anchor='center', text='HeadNum')
            tree_common.heading('Direction', anchor='center', text='Direction')
            tree_separate.heading('#0',text='')
            tree_separate.heading('Machine_No', anchor='center', text='Machine_No')
            tree_separate.heading('DXPCON Version', anchor='center', text='DXPCON Version')
            tree_separate.heading('DPC Version', anchor='center', text='DPC Version')
            tree_separate.heading('MATE Unit Type',anchor='center', text='MATE Unit Type')
            tree_separate.heading('MATE FPGA Board Revision', anchor='center', text='MATE FPGA Board Revision')
            tree_separate.heading('MATE DMD Board Revision', anchor='center', text='MATE DMD Board Revision')
            tree_separate.heading('MATE App Software Version', anchor='center', text='MATE App Software Version')
            tree_separate.heading('MATE Interface SoC Version', anchor='center', text='MATE Interface SoC Version')
            tree_separate.heading('MATE Plot FPGA Version', anchor='center', text='MATE Plot FPGA Version')
            tree_separate.heading('MATE Head FPGA Version', anchor='center', text='MATE Head FPGA Version')
            tree_separate.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
            tree_separate.heading('DMD Type', anchor='center', text='DMD Type')
            tree_separate.heading('DMD DLPC910 Version', anchor='center', text='DMD DLPC910 Version')
            tree_separate.heading('MATE Test Pattern Revision', anchor='center', text='MATE Test Pattern Revision')
            tree_separate.heading('MATE Area Mask Pattern Revision', anchor='center', text='MATE Area Mask Pattern Revision')
            tree_separate.heading('MATE MAC Address', anchor='center', text='MATE MAC Address')
            tree_separate.heading('Updater', anchor='center', text='Updater')
            tree_separate.heading('Updatedate', anchor='center', text='Updatedate')
            tree_separate.heading('Customer', anchor='center', text='Customer',command=lambda: treeview_sort_customercolumn(tree_common, tree_separate, 'Customer'))
            tree_separate.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for loop in info:
                if settei.get() == 0:
                    cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                else:
                    cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? AND latestflag = ?', (loop[0], 0, "Latest"))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    headsinfo = [['----']*16]
                cur.execute('SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dpcver = cur.fetchone()
                if dpcver:
                    print(f"DPC Information - ID: {dpcver[0]}, Version: {dpcver[1]}")
                else:
                    dpcver = ['----' for temp in range(2)] 
                cur.execute('SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?', (loop[0], "Latest"))
                dxpconver = cur.fetchone()
                if dxpconver:
                    print(f"DPC Information - ID: {dxpconver[0]}, Version: {dxpconver[1]}")
                else:
                    dxpconver = ['----' for temp in range(2)] 
                if headsinfo:
                    print("Heads Information:")
                    for head in headsinfo:
                        cur.execute(query_pe,(head[1], loop[0], loop[0], loop[0]))
                        updatedate = cur.fetchone()
                        if updatedate:
                            print(f"Updatedate - Date: {updatedate[0]}")  
                        else:
                            updatedate =  ['----']
                        tree_common.insert(parent='', index='end', iid=count, values=(loop[0], loop[1], head[1], head[2]))
                        tree_separate.insert(parent='', index='end', iid=count, values=(loop[3], dxpconver[1], dpcver[1], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], updatedate[0], updatedate[1], loop[2], ""))                     
                        count +=1
                else:
                    print(f"Machine with ID {loop[0]} does not exist.")
                    return  
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL)
            tree_separate.configure(xscrollcommand = scrollbarx.set)
            scrollbarx[ 'command' ] = tree_separate.xview
            scrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
            #マウスホイールの同期
            tree_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_separate.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_common.place(relheight=1.0,relwidth=0.52)
            tree_separate.place(relx=0.52, relheight=1.0, relwidth=0.48)
            # 縦スクロールバー
            scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL)
            tree_separate.configure(yscroll = scrollbary.set)
            scrollbary[ 'command' ] = sync_tree_separate_yview
            scrollbary.place(relx=0.985, relheight=0.98)
            save_button['state'] = 'normal'
            reload()
    else:
        print(f"No registration data available for machine type {machine_name}.")
        tmsg.showerror("Error",f"機種 {machine_name} の登録データがありません.")
    
#テーブルのデータ削除
def delete_data(conn,machine_id):
    if machine_id != '':
        try:
            cur=conn.cursor()
            cur.execute('SELECT * FROM  Machines WHERE id = ?', (machine_id,))
            info= cur.fetchone()
            if info:
                if info[4] == "LAIZA":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_laiza Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                elif info[4] == "LIVERA":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_livera Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                elif info[4] == "ILIA":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_ilia Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                elif info[4] == "LACS":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_lacs Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                elif info[4] == "PE" or info[4] == "PE-Ver2":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_pe Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                elif info[4] == "MATE" or info[4] == "MATE3":
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM Heads_mate Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
                else:
                    cur.execute('DELETE FROM Machines Where id = ?',(machine_id,))
                    cur.execute('DELETE FROM DPC Where machines_id = ?',(machine_id,))
                    cur.execute('DELETE FROM DXPCON Where machines_id = ?',(machine_id,))
                    conn.commit()
                    tmsg.showinfo("Complete","削除完了")
            else:
                print(f"Product ID {machine_id} does not exist.")
                tmsg.showerror("Error",f"製番 {machine_id} が存在しません")
                return

        except sqlite3.Error as e:
            print(f"An error occurred: {e}")
            tmsg.showerror("Error",f"An error occurred: {e}")
    else:
        tmsg.showerror("Error","製番が空白です")

#主テーブルの更新
def update_machine_id(conn, machine_id, machine_name, customer, modelnum, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        # プライマリキーの存在チェック
        cur.execute("SELECT * FROM Machines WHERE id = ? ", (machine_id,))
        temp = cur.fetchone()
        
        engine = ""
        if machine_name in laizalist:
            engine = "LAIZA"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name)
        elif machine_name in liveralist:
            engine = "LIVERA"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name)
        elif machine_name in ilialist:
            engine = "ILIA"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name)
        elif machine_name in lacslist:
            engine = "LACS"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name)  
        elif machine_name in pe2list:
            engine = "PE-Ver2"
            if engine in machine_name:
                machine_name = separate_name(machine_name)
        elif machine_name in pelist:
            engine = "PE"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name)
        elif machine_name in pelist:
            engine = "MATE"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name) 
        elif machine_name in pelist:
            engine = "MATE3"
            if engine in machine_name.upper():
                machine_name = separate_name(machine_name) 
        if temp:
            if customer_button_check.get() ==1 and machinetype_button_check.get() ==1 and modelnum_button_check.get() == 1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (machine_name, customer, modelnum, engine, loginname, updatedate, machine_id))
                conn.commit()
            elif customer_button_check.get() == 1 and machinetype_button_check.get() == 1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (machine_name, customer, temp[3], engine, loginname, updatedate, machine_id))
                conn.commit()
            elif customer_button_check.get() == 1 and modelnum_button_check.get() == 1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (temp[1], customer, modelnum, temp[4], loginname, updatedate, machine_id))
                conn.commit()
            elif machinetype_button_check.get() ==1 and modelnum_button_check.get() == 1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (machine_name, temp[2], modelnum, engine, loginname, updatedate, machine_id))
                conn.commit()
            elif customer_button_check.get() ==1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (temp[1], customer, temp[3],temp[4] , loginname, updatedate, machine_id))
                conn.commit()
            elif machinetype_button_check.get() ==1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (machine_name, temp[2], temp[3], engine, loginname, updatedate, machine_id ))
                conn.commit()
            elif modelnum_button_check.get() == 1:
                cur.execute('''UPDATE Machines SET name = ?, customer = ?, modelnum = ?, enginetype = ?, loginname = ?, updatedate = ? WHERE id = ?''', (temp[1], temp[2], modelnum, temp[4], loginname, updatedate, machine_id ))
                conn.commit()
            print("UPDATE successfully")
            msg = "UPDATE successfully"
            return True,msg
        else:
            # プライマリキーが存在しない場合、何もしない
            print("ProductID dose not exist")
            msg = "ProductID dose not exist"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg     
    
#主テーブルの確認
def check_machine_id(conn, machine_id):
    try:
        cur=conn.cursor()
        # プライマリキーの存在チェック
        cur.execute("SELECT 1 FROM Machines WHERE id = ? ", (machine_id,))
        if cur.fetchone() is None:
            # プライマリキーが存在しない場合、何もしない
            print("ProductID dose not exist")
            msg = "ProductID dose not exist"
            return False,msg
        else:
            msg = "ProductID exist"
            return True,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg  

#副テーブルの更新(laiza)
def update_heads_laiza(conn, machine_id, headnum, direction ,unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号を更新条件とする。
        cur.execute("SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? ", (machine_id, headnum, direction, "Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!=unit or info[4]!= vpb or info[5]!= hpb or info[6] != app or info[7] !=vpfpga or info[8] != ipfpga or info[9] != hpfpga or info[10] != ftp or info[11] != xtp
                        or info[12]!= amp or info[13]!= ma or info[14] != dmdunit or info[15] !=dmdtype or info[16]!= dmdversion):
                            cur.execute('''UPDATE Heads_laiza SET latestflag = 'Older'
                                        WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                                        ''', (machine_id, headnum, direction, "Latest"))
                            cur.execute("INSERT INTO Heads_laiza (machines_id, headnum, direction, unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                        ( machine_id, headnum, direction, unit, vpb, hpb, app, vpfpga, ipfpga, hpfpga, ftp, xtp, amp, ma, dmdunit, dmdtype, dmdversion, loginname, updatedate,"Latest"))
                            conn.commit()
                            msg = "UPDATE Heads_laiza Success"
                            return True,msg 
            else:
                msg = "Data update results remain the same."
                return True,msg 
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg
    
#副テーブルの更新(livera)
def update_heads_livera(conn, machine_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号を更新条件とする。
        cur.execute("SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? ", (machine_id, headnum, direction, "Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!=unit or info[4]!=br or info[5]!=bsv or info[6]!= app or info[7]!= fpga or info[8]!= ftp or info[9]!= xtp or info[10]!=amp or info[11]!= ma or info[12] != dmdunit 
               or info[13] != dmdbr or info[14] != dmdfpga or info[15] != dmdtype or info[16] != dmdversion):
                cur.execute('''UPDATE Heads_livera SET latestflag = 'Older'
                            WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                            ''', (machine_id, headnum, direction, "Latest"))
                cur.execute("INSERT INTO Heads_livera (machines_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                          ( machine_id, headnum, direction, unit, br, bsv, app, fpga, ftp, xtp, amp, ma, dmdunit, dmdbr, dmdfpga, dmdtype, dmdversion, loginname, updatedate, "Latest"))
                conn.commit()   
                msg = "UPDATE Heads_livera Success"
                return True,msg 
            else:
                msg = "Data update results remain the same."
                return True,msg 
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg

#副テーブルの更新(ilia)
def update_heads_ilia(conn, machine_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号を更新条件とする。
        cur.execute("SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ? AND direction= ? AND latestflag = ? ", (machine_id,headnum,direction,"Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!= unit or info[4] !=uibr or info[5] != hpbr or info[6] != app or info[7] != isv or info[8] != pfr or info[9] != dmdunit or info[10] != dmdtype 
               or info[11] != dmdddc or info[12] != tp or info[13] != amtp or info[14] != ma ):
                cur.execute('''UPDATE Heads_ilia SET latestflag = 'Older'
                            WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                            ''', (machine_id, headnum, direction, "Latest"))
                cur.execute("INSERT INTO Heads_ilia (machines_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                             ( machine_id, headnum, direction, unit, uibr, hpbr, app, isv, pfr, dmdunit, dmdtype, dmdddc, tp, amtp, ma, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE Heads_livera Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg  
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg  
    
#副テーブルの更新(lacs)
def update_heads_lacs(conn, machine_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号と最新版とバージョンが異なることを更新条件とする
        cur.execute("SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? ", (machine_id, headnum, direction, "Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!= unit or info[4] !=unirev or info[5] != bootsw or info[6] != softv or info[7] != fpgarev or info[8] !=ddhard or info[9] != xtp or info[10] != amp):
                cur.execute('''UPDATE Heads_lacs SET latestflag = 'Older'
                            WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                            ''', (machine_id, headnum, direction, "Latest"))
                cur.execute("INSERT INTO Heads_lacs (machines_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (  machine_id, headnum, direction, unit, unirev, bootsw, softv, fpgarev, ddhard, xtp, amp, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE Heads_lacs Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg 
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")

#副テーブルの更新(PE)
def update_heads_pe(conn, machine_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号と最新版のバージョンが異なることを更新の条件とする
        cur.execute("SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? ", (machine_id, headnum, direction, "Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!=bootsw or info[4]!=cpufpga or info[5]!=mccon or info[6]!= revs or info[7]!= mcdoft or info[8]!= hardrev or info[9] !=fpgahrev or info[10] != tpr):
                cur.execute('''UPDATE Heads_pe SET latestflag = 'Older'
                            WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                            ''', (machine_id, headnum, direction, "Latest"))
                cur.execute("INSERT INTO Heads_pe (machines_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (machine_id, headnum, direction, bootsw, cpufpga, mccon, revs, mcdoft, hardrev, fpgahrev, tpr, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE Heads_PE Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg 
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")

#副テーブルの更新(MATE)
def update_heads_mate(conn, machine_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDとヘッド番号と最新版のバージョンが異なることを更新の条件とする
        cur.execute("SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? ",(machine_id,headnum,direction,"Latest"))
        info = cur.fetchone()
        if info:
            if(info[3]!= ut or info[4]!=fpgabr or info[5]!= dmdbr or info[6]!=app or info[7]!=isoc or info[8]!=pfpga or info[9]!=hfpga or info[10]!=dmdunit or info[11]!=dlpcver or info[12]!=tp or info[13]!=amp or info[14]!=mc):
                cur.execute('''UPDATE Heads_mate SET latestflag = 'Older'
                            WHERE machines_id = ? AND headnum = ? AND direction = ? AND latestflag = ? 
                            ''', (machine_id, headnum, direction, "Latest"))
                cur.execute("INSERT INTO Heads_mate (machines_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (machine_id, headnum, direction, ut, fpgabr, dmdbr, app, isoc, pfpga, hfpga, dmdunit, dmdtype, dlpcver, tp, amp, mc, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE Heads_Mate Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg 
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")

#副テーブルの更新(DPC)
def update_dpc(conn, machine_id, dpcver, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDを更新条件とする
        cur.execute("SELECT * FROM DPC WHERE machines_id = ? AND latestflag = ?", (machine_id,"Latest"))
        info = cur.fetchone()
        if info:
            if(info[1]!= dpcver):
                cur.execute('''UPDATE DPC SET latestflag = 'Older'
                WHERE machines_id = ? AND latestflag = ? 
                ''', (machine_id, "Latest"))
                cur.execute("INSERT INTO DPC (machines_id, dataconver, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?)",
                            (machine_id, dpcver, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE DPC Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg    
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg
    
#副テーブルの更新(DXPCON)
def update_dxpcon(conn, machine_id, dxpconver, updatedate):
    global loginname
    try:
        cur=conn.cursor()
        #装置IDを更新条件とする
        cur.execute("SELECT * FROM DXPCON WHERE machines_id = ? AND latestflag = ?", (machine_id,"Latest"))
        info = cur.fetchone()
        if info:
            if(info[1]!= dxpconver):
                cur.execute('''UPDATE DXPCON SET latestflag = 'Older'
                WHERE machines_id = ? AND latestflag = ? 
                ''', (machine_id, "Latest"))
                cur.execute("INSERT INTO DXPCON (machines_id, dxpconver, loginname, updatedate, latestflag) VALUES (?, ?, ?, ?, ?)",
                            (machine_id, dxpconver, loginname, updatedate, "Latest"))
                conn.commit()
                msg = "UPDATE DXPCON Success"
                return True,msg 
            else:
                print("Does not meet the requirements")
                msg = "Does not meet the requirements"
                return False,msg
        else:
            print("Does not meet the requirements")
            msg = "Does not meet the requirements"
            return False,msg
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
        msg = f"An error occurred: {e}"
        return False,msg

def parse_line(line):
    # Extract the test value after the colon
    return line.split(":")[1].strip().split(",")[0]

#edac_PE用データの並び替え
def transform_result(result):
    # 最初のインデックスを取得し、対応する値を辞書に格納
    value_dict = {}
    for sublist in result:
        index = sublist[0]
        values = sublist[1:]
        value_dict[index] = values

    # 最初のサブリストの長さから1を引いた値をループ範囲に使用
    num_elements = len(result[0]) - 1

    transformed_list = []
    # 新しいリストを生成
    for i in range(num_elements):
        new_sublist = [i]
        new_sublist.append(value_dict[1][i])
        new_sublist.append(value_dict[2][i])
        new_sublist.append(value_dict[3][0])
        new_sublist.append(value_dict[4][0])
        new_sublist.append(value_dict[5][i])
        new_sublist.append(value_dict[6][0])
        new_sublist.append(value_dict[7][i])
        new_sublist.append(value_dict[8][i])
        transformed_list.append(new_sublist)

    return transformed_list

#PE用 確認のため
def confirm_pe(line):
    parts = line.split()
    pe = parts[1][0:2]
    return pe

#ファイルの読み込み・分解(EDAC)
def process_file_edac(file_path):
    list1 = []
    list2 = []
    current_sublist = []
    current_index = None
    count = 0
    count2 = 0
    peflag = False
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            if line.strip():  # 空白行の読み飛ばし
                if count == 0:
                    prefix = line.split()[1].split('_')[0]
                    try:
                        index = int(prefix[-1])
                    except ValueError:
                        continue
                if count == 0:
                    if 'Unit Type' in line:
                        peflag = False
                    else:
                        peflag = True                    
                count +=1
                if peflag != True:
                    prefix = line.split()[1].split('_')[0]
                    index = int(prefix[-1])
                    if current_index is None or current_index != index:
                        if current_sublist:
                            list1.append(current_sublist)
                        current_sublist = [index]
                        current_index = index
                    current_sublist.append(parse_line(line))
                else:
                    prefix = line.split()[1].split('_')[0]
                    try:
                        index = int(prefix[-1])
                    except ValueError:
                        index = 0
                    if index ==0:
                        count2+=1
                    if current_index is None or index ==0:
                        if current_sublist:
                            list1.append(current_sublist)
                        current_sublist = [count2]
                        current_index = count2
                    current_sublist.append(parse_line(line))
                    pe = confirm_pe(line)
                    if pe not in list2:
                        list2.append(pe)
        if current_sublist:
            list1.append(current_sublist)
    if peflag:
        list1 = transform_result(list1)
    return list1, list2

#ファイルの読み込み・分解(DPC)
def process_file_dpc(file_path):
    dpclist =[]
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            if line.strip():
                dpclist.append(line.split("Version")[1].strip().split("(")[0])
    return dpclist

#ファイルの読み込み・分解(DXPCON)
def process_file_dxpcon(file_path):
    dxplist =[]
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            if line.strip():
                dxplist.append(line.split("(")[1].rstrip('\n')[:-1]) 
    return dxplist

#ファイルのPathを取得するイベント
def get_textpath(event):
    ddtextbox.delete(0.,END)
    if event.data.startswith('{'):
        temp = event.data[1:-1]
        temp = temp.replace('\\','/')
        ddtextbox.insert(END,temp)
    else:
        ddtextbox.insert(END,event.data)

#フォルダ内のファイルリストの取得
def list_files_recursively(directory):
    list = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            list.append(os.path.join(root, file))
    return list

#ディレクトリで一番タイムスタンプが新しいファイルを取得
def file_get(directory_path):
    # リストディレクトリの内容
    files = os.listdir(directory_path)
    # フルパスを取得
    full_paths = [os.path.join(directory_path, f) for f in files]
    # 最も新しいファイルを見つける
    latest_file = max(full_paths, key=os.path.getmtime)
    return latest_file

# 結合されたセルの補正関数(客先だけ補正)
def fill_merged_cells(df):
    col = df.columns[1]
    last_val = None
    for i in range(len(df)):
        if pd.isnull(df.at[i, col]):
            df.at[i, col] = last_val
        else:
            last_val = df.at[i, col]
    return df

#文字列の分割
def separate_name(name):
    target = '('
    idx = name.find(target)
    return name[:idx]

#コンボボックスのlist追加
def append_list(name):
    global laizalist
    global liveralist
    global ilialist
    global lacslist
    global pe2list
    global pelist
    global matelist
    global mate3list
    global enginelist
    global machinelist
    with open(name, 'r', encoding='utf-8') as file:
       for line in file:
            #空白行の読み飛ばし
            if line != '\n':
                if '#' in line:
                    enginelist.append(line[1:].replace('\n',''))
                else:
                    if enginelist[-1] == "LAIZA":
                        laizalist.append(line.replace('\n',''))
                    elif enginelist[-1] == "LIVERA":
                        liveralist.append(line.replace('\n',''))
                    elif enginelist[-1] == "ILIA":
                        ilialist.append(line.replace('\n',''))
                    elif enginelist[-1] == "LACS":
                        lacslist.append(line.replace('\n',''))
                    elif enginelist[-1] == "PE":
                        pelist.append(line.replace('\n',''))
                    elif enginelist[-1] == "PE-Ver2":
                        pe2list.append(line.replace('\n',''))
                    elif enginelist[-1] == "MATE":
                        matelist.append(line.replace('\n',''))
                    elif enginelist[-1] == "MATE3":
                        mate3list.append(line.replace('\n',''))
    machinelist = laizalist + liveralist + ilialist + lacslist + pe2list +pelist + matelist + mate3list
               
#データの登録
def Click_reg_data(conn, path, id):
    try:
        if path != '':
            dt_now = datetime.now()  
            cur=conn.cursor()
            cur.execute('SELECT * FROM BasicInfo WHERE machines_id = ? AND enginetype IS NOT NULL', (id,))
            baseinfo = cur.fetchone()
            if baseinfo:
                insert_machine_id_if_not_exists(conn, id, baseinfo[1], baseinfo[2], baseinfo[3], baseinfo[4], dt_now.strftime('%Y/%m/%d %H:%M'))
                bsname = os.path.basename(path)
                if bsname == 'Version':
                    temp = list_files_recursively(path)
                    for i in range(len(temp)):
                        fullpath = temp[i].replace('\\','/')
                        tempname = os.path.basename(fullpath)
                        if tempname == 'edac.ver':
                            result_list1, result_list2 = process_file_edac(fullpath)
                            side = show_side_selection_dialog(root)
                            if side not in ['Left', 'Right']:
                                tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください。")
                            else :
                                if baseinfo[4] == 'LAIZA' and result_list1[0][1] == 'EDAC-LI':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_livera or reg_ilia or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4],result_list1[i][5],result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif baseinfo[4] == 'LIVERA' and result_list1[0][1] == 'EDAC-LV':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_ilia or reg_laiza or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif baseinfo[4] == 'ILIA' and result_list1[0][1] == 'EDAC-IA':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4],result_list1[i][5], result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")                                    
                                elif baseinfo[4] == 'LACS' and result_list1[0][1] == 'EDAC-LE':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_lacs(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5],
                                                                                result_list1[i][6], result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif baseinfo[4] == 'PE' and result_list2[0] == 'PE' or baseinfo[4] == 'PE-Ver2' and result_list2[0] == 'PE':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_pe(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                            result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif baseinfo[4] == 'MATE' and result_list1[0][1] == 'EDAC-MT' or baseinfo[4] == 'MATE3' and result_list1[0][1] == 'Undefined':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_pe:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_mate(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                else:
                                    tmsg.showerror("Error","マシンとエンジンの情報が一致しません")                      
                        elif tempname == 'datacon.ver':
                            result_list1 = process_file_dpc(fullpath)
                            for i in range(len(result_list1)):
                                result1,resultmsg = insert_dpc(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                                if result1 == False:
                                    tmsg.showerror("Error",f"{resultmsg}")
                                    break
                            if result1 == True:
                                tmsg.showinfo("Complete","DPC 登録完了")
                        elif tempname == 'dxpcon.ver':
                            result_list1 = process_file_dxpcon(fullpath)
                            for i in range(len(result_list1)):
                                result1,resultmsg = insert_dxpcon(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                                if result1 == False:
                                    tmsg.showerror("Error",f"{resultmsg}")
                                    break
                            if result1 == True:
                                tmsg.showinfo("Complete","DXPCON 登録完了")
                    dir_to_zip(path,id)
                elif bsname == 'edac.ver':
                    # カスタムダイアログで右・左を選択
                    side = show_side_selection_dialog(root)
                    if side not in [ 'Left','Right']:
                        tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください。")
                    else :
                        result_list1, result_list2 = process_file_edac(path)
                        if baseinfo[4] == 'LAIZA' and result_list1[0][1] == 'EDAC-LI':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_livera or reg_ilia or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        elif baseinfo[4] == 'LIVERA' and result_list1[0][1] == 'EDAC-LV':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_ilia or reg_laiza or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        elif baseinfo[4] == 'ILIA' and result_list1[0][1] == 'EDAC-IA':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4],result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        elif baseinfo[4] == 'LACS' and result_list1[0][1] == 'EDAC-LE':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_lacs(conn, id, result_list1[i][0], side,result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5],
                                                                            result_list1[i][6], result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        elif baseinfo[4] == 'PE' and result_list2[0] == 'PE' or baseinfo[4] == 'PE-Ver2' and result_list2[0] == 'PE':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_pe(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                        result_list1[i][7], result_list1[i][8],  dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        elif baseinfo[4] == 'MATE' and result_list1[0][1] == 'EDAC-MT' or baseinfo[4] == 'MATE3' and result_list1[0][1] == 'Undefined':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_pe:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_mate(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                        result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                                if result1 == True:
                                    file_to_zip(path,id)
                                    tmsg.showinfo("Complete","登録完了")
                        else:
                            tmsg.showerror("Error","Machine and engine information does not match")
                elif bsname == 'datacon.ver':
                    result_list1 = process_file_dpc(path)
                    for i in range(len(result_list1)):
                        result1,resultmsg = insert_dpc(conn, id, result_list1[i],dt_now.strftime('%Y/%m/%d %H:%M'))
                        if result1 == False:
                            tmsg.showerror("Error",f"{resultmsg}")
                            break
                    if result1 == True:
                        file_to_zip(path,id)
                        tmsg.showinfo("Complete","登録完了")
                elif bsname == 'dxpcon.ver':
                    result_list1 = process_file_dxpcon(path)
                    for i in range(len(result_list1)):
                        result1,resultmsg = insert_dxpcon(conn, id, result_list1[i],dt_now.strftime('%Y/%m/%d %H:%M'))
                        if result1 == False:
                            tmsg.showerror("Error",f"{resultmsg}")
                            break
                    if result1 == True:
                        file_to_zip(path,id)
                        tmsg.showinfo("Complete","登録完了")            
                else :
                    tmsg.showerror("Error","Incorrect file name")
            else:
                cur.execute('SELECT * FROM Machines WHERE id = ?', (id,))
                typeinfo = cur.fetchone()
                if typeinfo:
                    name = typeinfo[1]
                    customer = typeinfo[2]
                    modelnum = typeinfo[3]
                    engine = typeinfo[4]
                else:
                    reroad_customerlistall(cur)
                    customer, name , modelnum = show_customer_selection_dialog(root)
                    if name == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    engine = ""
                    if name in laizalist:
                        engine = "LAIZA"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in liveralist:
                        engine = "LIVERA"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in ilialist:
                        engine = "ILIA"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in lacslist:
                        engine = "LACS"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in pe2list:
                        engine = "PE-Ver2"
                        if engine in name:
                            name = separate_name(name)
                    elif name in pelist:
                        engine = "PE"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in matelist:
                        engine = "MATE"
                        if engine in name.upper():
                            name = separate_name(name)
                    elif name in mate3list:
                        engine = "MATE3"
                        if engine in name.upper():
                            name = separate_name(name)
                insert_machine_id_if_not_exists(conn,id,name,customer,modelnum,engine,dt_now.strftime('%Y/%m/%d %H:%M'))
                bsname = os.path.basename(path)
                if bsname == 'Version':
                    temp = list_files_recursively(path)
                    for i in range(len(temp)):
                        fullpath = temp[i].replace('\\','/')
                        tempname = os.path.basename(fullpath)
                        if tempname == 'edac.ver':
                            result_list1, result_list2 = process_file_edac(fullpath)
                            side = show_side_selection_dialog(root)
                            if side not in ['Left', 'Right']:
                                tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください。")
                            else :
                                if engine == "LAIZA" and result_list1[0][1] == 'EDAC-LI':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_livera or reg_ilia or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif engine == "LIVERA" and result_list1[0][1] == 'EDAC-LV':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_ilia or reg_laiza or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif engine == "ILIA" and result_list1[0][1] == 'EDAC-IA':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_lacs or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif engine == "LACS" and result_list1[0][1] == 'EDAC-LE':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_pe or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_lacs(conn, id, result_list1[i][0], side,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5],
                                                                                result_list1[i][6], result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了")
                                elif engine == 'PE' and result_list2[0] == 'PE' or engine == 'PE-Ver2' and result_list2[0] == 'PE':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                                    reg_mate = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_mate:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_pe(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                                result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","エンジン登録完了") 
                                elif engine == 'MATE' and result_list1[0][1] == 'EDAC-MT' or engine == 'MATE3' and result_list1[0][1] == 'Undefined':
                                    cur=conn.cursor()
                                    cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                                    reg_laiza =  cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                                    reg_livera = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                                    reg_ilia = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                                    reg_lacs = cur.fetchone()
                                    cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                                    reg_pe = cur.fetchone()
                                    if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_pe:
                                        tmsg.showerror("Error","製番は他のテーブルで使用されています")
                                    else:
                                        for i in range(len(result_list1)):
                                            result1,resultmsg = insert_heads_mate(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                                result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                                            if result1 == False:
                                                tmsg.showerror("Error",f"{resultmsg}")
                                                break
                                        if result1 == True:
                                            tmsg.showinfo("Complete","製番は他のテーブルで使用されています")
                                else:
                                    tmsg.showerror("Error","機種とエンジンの情報が一致しません")                        
                        elif tempname == 'datacon.ver':
                            result_list1 = process_file_dpc(fullpath)
                            for i in range(len(result_list1)):
                                result1,resultmsg = insert_dpc(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                                if result1 == False:
                                    tmsg.showerror("Error",f"{resultmsg}")
                                    break
                            if result1 == True:
                                tmsg.showinfo("Complete","DPC 登録完了")
                        elif tempname == 'dxpcon.ver':
                            result_list1 = process_file_dxpcon(fullpath)
                            for i in range(len(result_list1)):
                                result1,resultmsg = insert_dxpcon(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                                if result1 == False:
                                    tmsg.showerror("Error",f"{resultmsg}")
                                    break
                            if result1 == True:
                                tmsg.showinfo("Complete","DXPCON 登録完了")
                    dir_to_zip(path,id)
                elif bsname == 'edac.ver':
                    # カスタムダイアログで右・左を選択
                    side = show_side_selection_dialog(root)
                    if side not in [ 'Left','Right']:
                        tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください。")
                    else :
                        result_list1, result_list2 = process_file_edac(path)
                        if engine == "LAIZA" and result_list1[0][1] == 'EDAC-LI':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_livera or reg_ilia or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                        elif engine == "LIVERA" and result_list1[0][1] == 'EDAC-LV':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_ilia or reg_laiza or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                        elif engine == "ILIA" and result_list1[0][1] == 'EDAC-IA':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_lacs or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                        elif engine == "LACS" and result_list1[0][1] == 'EDAC-LE':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_pe or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_lacs(conn, id, result_list1[i][0], side,result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5],
                                                                            result_list1[i][6], result_list1[i][7], result_list1[i][8],  dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                        elif engine == "PE" and result_list2[0] == 'PE' or engine == "PE-Ver2" and result_list2[0] == 'PE':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_mate WHERE machines_id = ?", (id,))
                            reg_mate = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_mate:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_pe(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                        result_list1[i][7], result_list1[i][8],  dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break 
                        elif engine == 'MATE' and result_list1[0][1] == 'EDAC-MT' or engine == 'MATE3' and result_list1[0][1] == 'Undefined':
                            cur=conn.cursor()
                            cur.execute("SELECT 1 FROM Heads_laiza WHERE machines_id = ?", (id,))
                            reg_laiza =  cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_livera WHERE machines_id = ?", (id,))
                            reg_livera = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_ilia WHERE machines_id = ?", (id,))
                            reg_ilia = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_lacs WHERE machines_id = ?", (id,))
                            reg_lacs = cur.fetchone()
                            cur.execute("SELECT 1 FROM Heads_pe WHERE machines_id = ?", (id,))
                            reg_pe = cur.fetchone()
                            if reg_laiza or reg_livera or reg_ilia or reg_lacs or reg_pe:
                                tmsg.showerror("Error","製番は他のテーブルで使用されています")
                            else:
                                for i in range(len(result_list1)):
                                    result1,resultmsg = insert_heads_mate(conn, id, result_list1[i][0], side ,result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                                        result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                                    if result1 == False:
                                        tmsg.showerror("Error",f"{resultmsg}")
                                        break
                        else:
                            tmsg.showerror("Error","機種とエンジンの情報が一致しない")
                            return
                        if result1 == True:
                            file_to_zip(path,id)
                            tmsg.showinfo("Complete","登録完了")
                elif bsname == 'datacon.ver':
                    result_list1 = process_file_dpc(path)
                    for i in range(len(result_list1)):
                        result1,resultmsg = insert_dpc(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                        if result1 == False:
                            tmsg.showerror("Error",f"{resultmsg}")
                            break
                    if result1 == True:
                        file_to_zip(path,id)
                        tmsg.showinfo("Complete","登録完了")
                elif bsname == 'dxpcon.ver':
                    result_list1 = process_file_dxpcon(path)
                    for i in range(len(result_list1)):
                        result1,resultmsg = insert_dxpcon(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                        if result1 == False:
                            tmsg.showerror("Error",f"{resultmsg}")
                            break
                    if result1 == True:
                        file_to_zip(path,id)
                        tmsg.showinfo("Complete","登録完了")            
                else :
                    tmsg.showerror("Error","ファイル名が不適切です")
        else:
            tmsg.showerror("Error","製番かファイル名が空白です")
    except:
        tmsg.showerror("Error","データ登録エラー")
        
#データの更新
def Click_update_data(conn,path,id):
    try:
        if path != '': 
            dt_now = datetime.now()
            check,msg = check_machine_id(conn,id)
            reroad_customerlistall(conn.cursor())
            if check:
                if customer_button_check.get() ==1  and machinetype_button_check.get() ==1 and modelnum_button_check.get() == 1:
                    customer, machinetype ,modelnum = show_customer_selection_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, customer, modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif customer_button_check.get() == 1 and machinetype_button_check.get() == 1:
                    customer, machinetype = show_customerandmachinetype_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, customer, "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif customer_button_check.get() == 1 and modelnum_button_check.get() == 1:
                    customer, modelnum = show_customerandmodel_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", customer, modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif machinetype_button_check.get() == 1 and modelnum_button_check.get() == 1:
                    machinetype,modelnum = show_modelandmachinetype_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, "", modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif customer_button_check.get() == 1:
                    customer = show_customeronly_selection_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", customer, "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif machinetype_button_check.get() == 1:
                    machinetype = show_machinetype_selection_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, "", "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
                elif modelnum_button_check.get() == 1:
                    modelnum = show_modelnumonly_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", "", modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新成功")
            else:
                tmsg.showerror("Error",f"{msg}")
            bsname = os.path.basename(path)
            cur=conn.cursor()
            cur.execute('SELECT * FROM Machines WHERE id = ?', (id,))
            typeinfo = cur.fetchone()
            name = typeinfo[4]
            if bsname == 'Version':
                temp = list_files_recursively(path)
                for i in range(len(temp)):
                    fullpath = temp[i].replace('\\','/')
                    tempname = os.path.basename(fullpath)
                    if tempname == 'edac.ver':
                        result_list1, result_list2 = process_file_edac(fullpath)
                        side = show_side_selection_dialog(root)
                        if side not in [ 'Left','Right']:
                            tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください")
                        else:
                            result2 = False
                            if name == 'LAIZA' and result_list1[0][1] == 'EDAC-LI':
                                for i in range(len(result_list1)):
                                    result2,msg2=update_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                            elif name == 'LIVERA' and result_list1[0][1] == 'EDAC-LV':
                                for i in range(len(result_list1)):
                                    result2,msg2=update_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                            elif name == 'ILIA' and result_list1[0][1] == 'EDAC-IA':
                                for i in range(len(result_list1)):
                                    result2,msg2=update_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                            elif name in 'LACS' and result_list1[0][1] == 'EDAC-LE':
                                for i in range(len(result_list1)):
                                    result2,msg2 = update_heads_lacs(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                            elif name == 'PE' and result_list2[0] == 'PE' or name == 'PE-Ver2' and result_list2[0] == 'PE':
                                for i in range(len(result_list1)):
                                    result2,msg2 = update_heads_pe(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                            elif name == 'MATE' and result_list1[0][1] == 'EDAC-MT' or name == 'MATE3' and result_list1[0][1] == 'Undefined':
                                for i in range(len(result_list1)):
                                    result2,msg2 = update_heads_mate(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2],result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                            result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                            else:
                                tmsg.showerror("Error","機種とエンジンの情報が一致しない")
                            if result2 == True:
                                tmsg.showinfo("Complete","エンジン情報の更新完了")
                            else:
                                tmsg.showinfo("Error","更新条件を満たしませんでした")
                    elif tempname == 'datacon.ver':
                        result_list1 = process_file_dpc(fullpath)
                        for i in range(len(result_list1)):
                            result1,resultmsg = update_dpc(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                            if result1 == False:
                                tmsg.showerror("Error",f"{resultmsg}")
                                break
                        if result1 == True:
                            tmsg.showinfo("Complete","DPC 情報の更新完了")
                    elif tempname == 'dxpcon.ver':
                        result_list1 = process_file_dxpcon(fullpath)
                        for i in range(len(result_list1)):
                            result1,resultmsg = update_dxpcon(conn, id, result_list1[i], dt_now.strftime('%Y/%m/%d %H:%M'))
                            if result1 == False:
                                tmsg.showerror("Error",f"{resultmsg}")
                                break
                        if result1 == True:
                            tmsg.showinfo("Complete","DXPCON 情報の更新完了")
                dir_to_zip(path,id)                   
            elif bsname == 'edac.ver':
                result_list1, result_list2 = process_file_edac(path)
                side = show_side_selection_dialog(root)
                if side not in [ 'Left','Right']:
                    tmsg.showerror("Error", "サイドの入力が無効です。「Left」または「Right」を選択してください")
                else:
                    result2 = False
                    if name == 'LAIZA' and result_list1[0][1] == 'EDAC-LI':
                        for i in range(len(result_list1)):
                            result2,msg2=update_heads_laiza(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                    elif name == 'LIVERA' and result_list1[0][1] == 'EDAC-LV':
                        for i in range(len(result_list1)):
                            result2,msg2=update_heads_livera(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], result_list1[i][14], dt_now.strftime('%Y/%m/%d %H:%M'))
                    elif name == 'ILIA' and result_list1[0][1] == 'EDAC-IA':
                        for i in range(len(result_list1)):
                            result2,msg2=update_heads_ilia(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], dt_now.strftime('%Y/%m/%d %H:%M'))
                    elif name == 'LACS' and result_list1[0][1] == 'EDAC-LE':
                        for i in range(len(result_list1)):
                            result2,msg2 = update_heads_lacs(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                    elif name == 'PE' and result_list2[0] == 'PE' or name == 'PE-Ver2' and result_list2[0] == 'PE':
                        for i in range(len(result_list1)):
                            result2,msg2 = update_heads_pe(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4],result_list1[i][5],result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], dt_now.strftime('%Y/%m/%d %H:%M'))
                    elif name == 'MATE' and result_list1[0][1] == 'EDAC-MT' or name == 'MATE3' and  result_list1[0][1] == 'Undefined':
                        for i in range(len(result_list1)):
                            result2,msg2=update_heads_mate(conn, id, result_list1[i][0], side, result_list1[i][1], result_list1[i][2], result_list1[i][3], result_list1[i][4], result_list1[i][5], result_list1[i][6],
                                    result_list1[i][7], result_list1[i][8], result_list1[i][9], result_list1[i][10], result_list1[i][11], result_list1[i][12], result_list1[i][13], dt_now.strftime('%Y/%m/%d %H:%M'))
                    else:
                        tmsg.showerror("Error","機種とエンジンの情報が一致しない")
                    if result2 == True:
                        tmsg.showinfo("Complete","更新完了")
                    else:
                        tmsg.showinfo("Error","更新条件を満たしませんでした")
                    file_to_zip(path,id)
            elif bsname == 'datacon.ver':
                result_list1 = process_file_dpc(path)
                for i in range(len(result_list1)):
                    result1,resultmsg = update_dpc(conn, id, result_list1[i],dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{resultmsg}")
                        break
                if result1 == True:
                    tmsg.showinfo("Complete","更新完了")
                else:
                    tmsg.showinfo("Error","更新条件を満たしませんでした")
                file_to_zip(path,id)
            elif bsname == 'dxpcon.ver':
                result_list1 = process_file_dxpcon(path)
                for i in range(len(result_list1)):
                    result1,resultmsg = update_dxpcon(conn, id, result_list1[i],dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{resultmsg}")
                        break
                if result1 == True:
                    tmsg.showinfo("Complete","更新完了")
                else:
                    tmsg.showinfo("Error","更新条件を満たしませんでした")
                file_to_zip(path,id)
            else:
                tmsg.showerror("Error","ファイル名が不適切です")
        else:
            dt_now = datetime.now()
            check,msg = check_machine_id(conn,id)
            reroad_customerlistall(conn.cursor())
            if check:
                if customer_button_check.get() ==1  and machinetype_button_check.get() ==1 and modelnum_button_check.get() == 1:
                    customer, machinetype ,modelnum = show_customer_selection_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, customer, modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif customer_button_check.get() == 1 and machinetype_button_check.get() == 1:
                    customer, machinetype = show_customerandmachinetype_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, customer, "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif customer_button_check.get() == 1 and modelnum_button_check.get() == 1:
                    customer, modelnum = show_customerandmodel_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", customer, modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif machinetype_button_check.get() == 1 and modelnum_button_check.get() == 1:
                    machinetype,modelnum = show_modelandmachinetype_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, "", modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif customer_button_check.get() == 1:
                    customer = show_customeronly_selection_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", customer, "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif machinetype_button_check.get() == 1:
                    machinetype = show_machinetype_selection_dialog(root)
                    if machinetype == "":
                        tmsg.showerror("Error", "ダイアログ入力エラー。機種を選択し、OKをクリックしてください。")
                        return
                    result1,msg = update_machine_id(conn, id, machinetype, "", "", dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                elif modelnum_button_check.get() == 1:
                    modelnum = show_modelnumonly_dialog(root)
                    result1,msg = update_machine_id(conn, id, "", "", modelnum, dt_now.strftime('%Y/%m/%d %H:%M'))
                    if result1 == False:
                        tmsg.showerror("Error",f"{msg}")
                    else:
                        tmsg.showinfo("Complete","メインの更新完了")
                else:
                    tmsg.showerror("Error","製番またはファイルが空白です")
            else:
                tmsg.showerror("Error",f"{msg}") 
    except:
        tmsg.showerror("Error","データ更新エラー")   

#ファイルの保存
def save_csv():
    global tree_common
    global tree_separate
    dt_now = datetime.now()
    defname = "DiVersion"
    csvname = defname + dt_now.strftime('%Y%m%d%H%M%S') + ".csv"
    if tree_common != None and tree_separate != None:
        # カラム識別子を取得
        column_ids_common = tree_common["columns"]
        column_ids_separete = tree_separate["displaycolumns"]
        # 各カラム識別子に対応するカラムヘッディングを取得
        column_common_headings = [tree_common.heading(col)["text"] for col in column_ids_common]
        column_separete_headings = [tree_separate.heading(col)["text"] for col in column_ids_separete]
        # treeウィジェットからデータを取得
        data_common = [tree_common.item(item)['values'] for item in tree_common.get_children()]
        data_separete = [[tree_separate.set(item, col) for col in column_ids_separete] for item in tree_separate.get_children()]
        # DataFrameを作成してデータを連結
        df_common = pd.DataFrame(data_common, columns=column_common_headings)
        df_separete = pd.DataFrame(data_separete, columns=column_separete_headings)
        df = pd.concat([df_common, df_separete], axis=1)
        # ファイルダイアログを表示して保存先を選択
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')], initialfile=csvname, title='CSVファイルの保存先を選択してください')

        # 保存先が選択された場合のみファイルを保存
        if file_path:
            df.to_csv(file_path, encoding='utf-8_sig', index=False)
            print(f"ファイルが保存されました: {file_path}")
            tmsg.showinfo("Complete","csv 出力完了")
        else:
            print("ファイル保存がキャンセルされました")
            # DataFrameをCSVファイルに保存
            tmsg.showinfo("Cancel","csv 出力がキャンセルされました")
    elif tree_common!= None:
        # カラム識別子を取得
        column_ids_common = tree_common["columns"]
        # 各カラム識別子に対応するカラムヘッディングを取得
        column_common_headings = [tree_common.heading(col)["text"] for col in column_ids_common]
        # treeウィジェットからデータを取得
        data_common = [tree_common.item(item)['values'] for item in tree_common.get_children()]
        # DataFrameを作成してデータを連結
        df_common = pd.DataFrame(data_common, columns=column_common_headings)
        # ファイルダイアログを表示して保存先を選択
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')], initialfile=csvname, title='CSVファイルの保存先を選択してください')
        # 保存先が選択された場合のみファイルを保存
        if file_path:
            df_common.to_csv(file_path, encoding='utf-8_sig', index=False)
            print(f"ファイルが保存されました: {file_path}")
            tmsg.showinfo("Complete","csv 出力完了")
        else:
            print("ファイル保存がキャンセルされました")
            # DataFrameをCSVファイルに保存
            tmsg.showinfo("Cancel","csv 出力がキャンセルされました") 
    else:
        tmsg.showerror("Error","データが表示されていません...")
 
#ファイルの保存(履歴表示)       
def logsave_csv():
    global tree_log_common
    global tree_log_separete
    dt_now = datetime.now()
    defname = "DiVersion"
    if tree_log_common != None and tree_log_separete != None:
        # カラム識別子を取得
        column_ids_common = tree_log_common["columns"]
        column_ids_separete = tree_log_separete["columns"]
        # 各カラム識別子に対応するカラムヘッディングを取得
        column_common_headings = [tree_log_common.heading(col)["text"] for col in column_ids_common]
        column_separete_headings = [tree_log_separete.heading(col)["text"] for col in column_ids_separete]
        # treeウィジェットからデータを取得
        data_common = [tree_log_common.item(item)['values'] for item in tree_log_common.get_children()]
        data_separete = [tree_log_separete.item(item)['values'] for item in tree_log_separete.get_children()]
        # DataFrameを作成してデータを連結
        df_common = pd.DataFrame(data_common, columns=column_common_headings)
        df_separete = pd.DataFrame(data_separete, columns=column_separete_headings)
        df = pd.concat([df_common, df_separete], axis=1)
        csvname = defname + dt_now.strftime('%Y%m%d%H%M%S') + ".csv"
        # ファイルダイアログを表示して保存先を選択
        file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')], initialfile=csvname, title='CSVファイルの保存先を選択してください')

        # 保存先が選択された場合のみファイルを保存
        if file_path:
            df.to_csv(file_path, encoding='utf-8_sig', index=False)
            print(f"ファイルが保存されました: {file_path}")
            tmsg.showinfo("Complete","csv 出力完了.")
        else:
            print("ファイル保存がキャンセルされました")
            # DataFrameをCSVファイルに保存
            tmsg.showinfo("Cancel","csv 出力がキャンセルされました")
    else:
        tmsg.showerror("Error","データが表示されていません...")

#Product_id制限
def at_least_char(string):
    global id_flag
    global authority
    #英数字のみ
    # alnumReg = re.compile(r'^[a-zA-Z0-9]+$')
    #数字のみ
    alnumReg = re.compile(r'^[0-9]+$')
    if alnumReg.match(string) or string == "":
        if len(string) < 5:
            if searchsettei.get() == 1:
                search_button['state'] = 'normal'
            else:
                search_button['state'] = 'disabled'
            register_button['state'] = 'disabled'
            update_botton['state'] = 'disabled'
            delete_button['state'] = 'disabled'
            logsearch_button['state'] = 'disabled' 
            id_flag = False         
        else:
            if authority == True:
                register_button['state'] = 'normal'
                update_botton['state'] = 'normal'
                delete_button['state'] = 'normal'
            search_button['state'] = 'normal'
            logsearch_button['state'] = 'normal'
            id_flag = True
        return len(string) <= 6
    else:
        return False

#Product_id制限
def at_least_charlog(string):
    #英数字のみ
    # alnumReg = re.compile(r'^[a-zA-Z0-9]+$')
    #数字のみ
    alnumReg = re.compile(r'^[0-9]+$')
    if alnumReg.match(string) or string == "":
        if len(string) < 5:
            logsearch_button['state'] = 'disabled'          
        else:
            logsearch_button['state'] = 'normal'
        return len(string) <= 6
    else:
        return False
    
#ラジオボタンのイベント
def on_id():
    global id_flag
    if id_flag == True:
        search_button['state'] = 'normal' 
    else:
        search_button['state'] = 'disabled' 
    
def on_cus():
    search_button['state'] = 'normal' 

def on_entry_text1_change(*args):
    entry_textlog_var.set(entry_text1_var.get())

#treeviewに表示するCalumnの選択
def reload():
    global tree_separate 
    if tree_separate != None:
        column_ids_separete = tree_separate["columns"]
        display_columns = []
        # 列の表示にあわせた表示列を作る
        if machine_no_button_check.get() == 1:
            display_columns.append(column_ids_separete[0])
        if dxpcon_button_check.get() == 1:
            display_columns.append(column_ids_separete[1])
        if dpc_button_check.get() == 1:
            display_columns.append(column_ids_separete[2])
        if engine_button_check.get() == 1:
            for num in column_ids_separete[3:-4]:
              display_columns.append(num)
        display_columns.append(column_ids_separete[-2])
        display_columns.append(column_ids_separete[-4])
        display_columns.append(column_ids_separete[-3])
        display_columns.append(column_ids_separete[-1]) 
        tree_separate["displaycolumns"] = display_columns

#ID一覧の取得
def display_idlist(conn):
    global tree_common
    global tree_separate
    global scrollbarx
    global scrollbary
    global sort_id
    if scrollbarx != None:
        scrollbarx.pack_forget()
        scrollbarx = None    
    if scrollbary != None:
        scrollbary.pack_forget()
        scrollbary = None
    if tree_common != None:
        tree_common.destroy()
        tree_common = None  # Treeviewを削除したことを示すためにNoneに設定
    if tree_separate != None:
        tree_separate.destroy()
        tree_separate = None  # Treeviewを削除したことを示すためにNoneに設定
    tree_common = ttk.Treeview(mainframe1,columns=main_column)
    tree_common.bind("<<TreeviewSelect>>", select_record)
    #列の設定
    tree_common.column('#0',width=0, stretch='no')
    tree_common.column('ID', anchor='center', width=200, stretch='no')
    tree_common.column('Model', anchor='center', width=200, stretch='no')
    tree_common.column('Customer', anchor='center', width=250, stretch='no')
    tree_common.column('Machine_No', anchor='center', width=200, stretch='no')
    tree_common.column('Engine', anchor='center', width=200, stretch='no')
    tree_common.column('blank', anchor='center', width=100, stretch='no')   
    #列の見出し
    tree_common.heading('#0',text='')
    tree_common.heading('ID', anchor='center', text='ID',command=lambda _col='ID': \
                     treeview_main_sort_column(tree_common, _col, False))
    tree_common.heading('Model', anchor='center', text='Model',command=lambda _col='Model': \
                     treeview_main_sort_column(tree_common, _col, False))
    tree_common.heading('Customer', anchor='center', text='Customer',command=lambda _col='Customer': \
                     treeview_main_sort_column(tree_common, _col, False))
    tree_common.heading('Machine_No', anchor='center', text='Machine_No',command=lambda _col='Machine_No': \
                     treeview_main_sort_column(tree_common, _col, False))
    tree_common.heading('Engine', anchor='center', text='Engine',command=lambda _col='Engine': \
                     treeview_main_sort_column(tree_common, _col, False))
    tree_common.heading('blank', anchor='center',text='')  
    cur=conn.cursor()
    cur.execute('SELECT * FROM  Machines ORDER BY id ASC')
    result = cur.fetchall()
    count = 0
    tree_common.tag_configure("red", foreground='red')
    if result:
        for info in result:
            tree_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], info[2], info[3], info[4], ""))
            count +=1
    else:
        print("infomation nothing")
    cur.execute('''SELECT * From BasicInfo WHERE machines_id NOT IN (SELECT id From Machines) ORDER BY machines_id ASC''')
    subresult = cur.fetchall()
    if subresult:
        for subinfo in subresult:
            tree_common.insert(parent='', index='end', iid=count, values=(subinfo[0], subinfo[1], subinfo[2], subinfo[3], subinfo[4], ""), tags="red")
            count +=1
    else:
        print("subinfomation nothing")
    
    # Styleの設定
    style_id = ttk.Style()
    style_id.map('Treeview',
                 foreground=[('selected', 'white')],
                 background=[('selected', 'deepskyblue')])
    style_id.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
    style_id.configure("Treeview.Heading", font=(None, 12))
    # Treeviewの枠線を非表示にする
    style_id.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])

    # 縦スクロールバー
    scrollbary = ttk.Scrollbar(mainframe1, orient=VERTICAL, command=tree_common.yview)
    tree_common.configure(yscroll = scrollbary.set)
    scrollbary.place(relx=0.986, relheight=0.98)
    #横スクロールバー
    scrollbarx = ttk.Scrollbar(mainframe1, orient=HORIZONTAL, command=tree_common.xview)
    tree_common.configure(xscroll = scrollbarx.set)
    scrollbarx.place(rely=0.98, relwidth=0.986)
    save_button['state'] = 'normal'
    # ウィジェットの配置
    tree_common.place(relheight=1.0,relwidth=1.0)
    
#装置id・装置名・エンジン情報の表示
def displaylog_machine_info(conn,machine_id):
    # Machinesテーブルから装置情報を取得
    if machine_id != '':
        global sort_ascending
        sort_ascending = False
        global sort_direction
        sort_direction = False
        global sort_time
        sort_time = True
        cur=conn.cursor()
        cur.execute('SELECT * FROM  Machines WHERE id = ?', (machine_id,))
        info= cur.fetchone()
        if info:
            print(f"Machine Information - ID: {info[0]}, Name: {info[1]}")
        else:
            print(f"ProductID {machine_id} does not exist.")
            tmsg.showerror("Error",f"製番 {machine_id} が存在しません")
            return
        global tree_log_common
        global tree_log_separete
        global logscrollbarx
        global logscrollbary
        if logscrollbarx != None:
            logscrollbarx.pack_forget()
            logscrollbarx = None   
        if logscrollbary != None:
            logscrollbary.pack_forget()
            logscrollbary = None 
        if tree_log_separete != None:
            tree_log_separete.destroy()
            tree_log_separete = None  # Treeviewを削除したことを示すためにNoneに設定
        if tree_log_common != None:
            tree_log_common.destroy()
            tree_log_common = None  # Treeviewを削除したことを示すためにNoneに設定
        if logsettei.get() == 0:
            cur.execute('SELECT * FROM DXPCON WHERE machines_id = ?', (machine_id,))
            dxpconinfo = cur.fetchall()
            if dxpconinfo:
                print("DXPCON Information exsit:")
            else:
                tmsg.showerror("Error","DXPCON の情報は登録されていません")
                return
            tree_log_common = ttk.Treeview(logframe1,columns=common_column_log)
            tree_log_separete = ttk.Treeview(logframe1,columns=DXPCON_colmn_log)
            #列の設定
            tree_log_common.column('#0', width=0, stretch='no')
            tree_log_common.column('ID', anchor='center', width=150, stretch='no')
            tree_log_common.column('Model', anchor='center', width=150, stretch='no')
            tree_log_separete.column('#0', width=0, stretch='no')
            tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
            tree_log_separete.column('DXPCON Version',anchor='center', width=200, stretch='no')
            tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
            tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
            tree_log_separete.column('blank', anchor='center', width=200, stretch='no')
            #列の見出し
            tree_log_common.heading('#0', text='')
            tree_log_common.heading('ID', anchor='center', text='ID')
            tree_log_common.heading('Model', anchor='center', text='Model')
            tree_log_separete.heading('#0', text='')
            tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
            tree_log_separete.heading('DXPCON Version',anchor='center', text='DXPCON Version')
            tree_log_separete.heading('Updater',anchor='center', text='Updater')
            tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate', command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
            tree_log_separete.heading('Customer',anchor='center', text='Customer')
            tree_log_separete.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for dxp in dxpconinfo:
                tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1]))
                tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], dxp[1], dxp[2], dxp[3], info[2], ""))
                count +=1
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
            tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
            logscrollbarx[ 'command' ] = tree_log_separete.xview
            logscrollbarx.place(relx=0.3125, rely=0.981, relwidth=0.675)
            #マウスホイールの同期
            tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_log_common.place(relwidth=0.3125, relheight=1.0)
            tree_log_separete.place(relx=0.3125, relwidth=0.6875,relheight=1.0)
            # 縦スクロールバー
            logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
            tree_log_separete.configure(yscroll = logscrollbary.set)
            logscrollbary[ 'command' ] = sync_tree_log_separete_yview
            logscrollbary.place(relx=0.986,relheight=0.982)
        elif logsettei.get() == 1:
            cur.execute('SELECT * FROM DPC WHERE machines_id = ?', (machine_id,))
            dataconinfo = cur.fetchall()
            if dataconinfo:
                print("DPC Information exsit:")
            else:
                tmsg.showerror("Error","DPC の情報は登録されていません.")
                return
            tree_log_common = ttk.Treeview(logframe1, columns=common_column_log)
            tree_log_separete = ttk.Treeview(logframe1, columns=DPC_colimn_log)
            #列の設定
            tree_log_common.column('#0', width=0, stretch='no')
            tree_log_common.column('ID', anchor='center', width=150, stretch='no')
            tree_log_common.column('Model', anchor='center', width=150, stretch='no')
            tree_log_separete.column('#0', width=0, stretch='no')
            tree_log_separete.column('Machine_No',anchor='center', width=220, stretch='no')
            tree_log_separete.column('DPC Version',anchor='center', width=220, stretch='no')
            tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
            tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
            tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
            tree_log_separete.column('blank', anchor='center', width=200, stretch='no')
            #列の見出し
            tree_log_common.heading('#0', text='')
            tree_log_common.heading('ID', anchor='center', text='ID')
            tree_log_common.heading('Model', anchor='center', text='Model')
            tree_log_separete.heading('#0', text='')
            tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
            tree_log_separete.heading('DPC Version',anchor='center', text='DXPCON Version')
            tree_log_separete.heading('Updater',anchor='center', text='Updater')
            tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate', command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
            tree_log_separete.heading('Customer',anchor='center', text='Customer')
            tree_log_separete.heading('blank', anchor='center', text='')
            #レコードの追加
            count =0
            for dpcver in dataconinfo:
                tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1]))
                tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], dpcver[1], dpcver[2], dpcver[3], info[2], ""))
                count +=1
            # Styleの設定
            style = ttk.Style()
            # Treeviewの選択時の背景色をデフォルトと同じにする
            style.map('Treeview', 
                    background=[('selected', style.lookup('Treeview', 'background'))],
                    foreground=[('selected', style.lookup('Treeview', 'foreground'))])
            style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
            style.configure("Treeview.Heading", font=(None, 12))
            # Treeviewの枠線を非表示にする
            style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
            #スクロールバーの追加
            logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
            tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
            logscrollbarx[ 'command' ] = tree_log_separete.xview
            logscrollbarx.place(relx=0.3125, rely=0.981, relwidth=0.675)
            #マウスホイールの同期
            tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
            tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
            # ウィジェットの配置
            tree_log_common.place(relwidth=0.3125, relheight=1.0)
            tree_log_separete.place(relx=0.3125, relwidth=0.6875,relheight=1.0)
            # 縦スクロールバー
            logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
            tree_log_separete.configure(yscroll = logscrollbary.set)
            logscrollbary[ 'command' ] = sync_tree_log_separete_yview
            logscrollbary.place(relx=0.986,relheight=0.982)
        elif logsettei.get() == 2:
            dire,format = show_engine_selection_dialog(root)
            if dire not in [ 'Both','Left','Right'] or format not in [ 'All','Head0']:
                tmsg.showerror("Error", "方向またはフォーマットの入力が無効です。")
            else:
                if info[4] == 'LAIZA':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC,', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_laiza WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                elif info[4] == 'LIVERA':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id,'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_livera WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                elif info[4] == 'ILIA':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_ilia WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                elif info[4] == 'LACS':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_lacs WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                elif info[4] == 'PE' or info[4] == 'PE-Ver2':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id,'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_pe WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                elif info[4] == 'MATE' or info[4] == 'MATE3':
                    if dire == 'Both':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? ORDER BY headnum ASC, direction ASC', (machine_id,))
                        else :
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ?', (machine_id, 0))
                    elif dire == 'Left':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Left'))
                        else :
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Left'))
                    elif dire == 'Right':
                        if format == 'All':
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND direction = ? ORDER BY headnum ASC', (machine_id, 'Right'))
                        else :
                            cur.execute('SELECT * FROM Heads_mate WHERE machines_id = ? AND headnum = ? AND direction = ?', (machine_id, 0, 'Right'))
                headsinfo = cur.fetchall()
                if headsinfo:
                    print("Heads Information exsit:")
                else:
                    tmsg.showerror("Error","エンジン情報は登録されていません")
                    return
                if info[4] == "LAIZA":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=laiza_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no')
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Unit Type',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Vector Process Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Head Process Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI App Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Vector Process FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Intersection Process FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Head Process FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI FHD Test Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI XGA Test Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI Area Mask Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LI MAC Address', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Unit Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('LI Unit Type',anchor='center', text='LI Unit Type')
                    tree_log_separete.heading('LI Vector Process Board Revision', anchor='center', text='LI Vector Process Board Revision')
                    tree_log_separete.heading('LI Head Process Board Revision', anchor='center', text='LI Head Process Board Revision')
                    tree_log_separete.heading('LI App Software Version', anchor='center', text='LI App Software Version')
                    tree_log_separete.heading('LI Vector Process FPGA Version', anchor='center', text='LI Vector Process FPGA Version')
                    tree_log_separete.heading('LI Intersection Process FPGA Version', anchor='center',text='LI Intersection Process FPGA Version')
                    tree_log_separete.heading('LI Head Process FPGA Version', anchor='center', text='LI Head Process FPGA Version')
                    tree_log_separete.heading('LI FHD Test Pattern Version', anchor='center', text='LI FHD Test Pattern Version')
                    tree_log_separete.heading('LI XGA Test Pattern Version', anchor='center', text='LI XGA Test Pattern Version')
                    tree_log_separete.heading('LI Area Mask Pattern Version', anchor='center',text='LI Area Mask Pattern Version')
                    tree_log_separete.heading('LI MAC Address', anchor='center', text='LI MAC Address')
                    tree_log_separete.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
                    tree_log_separete.heading('DMD Type', anchor='center', text='DMD Type')
                    tree_log_separete.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo:
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], head[17], head[18], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.98)
                elif info[4] == "LIVERA":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=livera_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no')
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV Unit Type',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV Boot Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV App Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV FHD Test Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV XGA Test Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV Area Mask Pattern Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LV MAC Address', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Unit Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD DDC4100 Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='Product ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('LV Unit Type',anchor='center', text='LV Unit Type')
                    tree_log_separete.heading('LV Board Revision', anchor='center', text='LV Board Revision')
                    tree_log_separete.heading('LV Boot Software Version', anchor='center', text='LV Boot Software Version')
                    tree_log_separete.heading('LV App Software Version', anchor='center', text='LV App Software Version')
                    tree_log_separete.heading('LV FPGA Version', anchor='center', text='LV FPGA Version')
                    tree_log_separete.heading('LV FHD Test Pattern Version', anchor='center',text='LV FHD Test Pattern Version')
                    tree_log_separete.heading('LV XGA Test Pattern Version', anchor='center', text='LV XGA Test Pattern Version')
                    tree_log_separete.heading('LV Area Mask Pattern Version', anchor='center', text='LV Area Mask Pattern Version')
                    tree_log_separete.heading('LV MAC Address', anchor='center', text='LV MAC Address')
                    tree_log_separete.heading('DMD Unit Type', anchor='center',text='DMD Unit Type')
                    tree_log_separete.heading('DMD Board Revision', anchor='center', text='DMD Board Revisions')
                    tree_log_separete.heading('DMD FPGA Version', anchor='center', text='DMD FPGA Version')
                    tree_log_separete.heading('DMD Type', anchor='center', text='DMD Type')
                    tree_log_separete.heading('DMD DDC4100 Version', anchor='center', text='DMD DDC4100 Version')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo:
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], head[17], head[18], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.98)
                elif info[4] == "ILIA":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=ilia_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no')
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no') 
                    tree_log_separete.column('IA Unit Type',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA USB Interface Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA Head Process Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA App Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA Interface SoC Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA Plot FPGA Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Unit Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD DDC4100', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA Test Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IA MAC Address', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='Product ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('IA Unit Type',anchor='center', text='IA Unit Type')
                    tree_log_separete.heading('IA USB Interface Board Revision', anchor='center', text='IA USB Interface Board Revision')
                    tree_log_separete.heading('IA Head Process Board Revision', anchor='center', text='IA Head Process Board Revision')
                    tree_log_separete.heading('IA App Software Version', anchor='center', text='IA App Software Version')
                    tree_log_separete.heading('IA Interface SoC Version', anchor='center', text='IA Interface SoC Version')
                    tree_log_separete.heading('IA Plot FPGA Revision', anchor='center',text='IA Plot FPGA Revision')
                    tree_log_separete.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
                    tree_log_separete.heading('DMD Type', anchor='center', text='DMD Type')
                    tree_log_separete.heading('DMD DDC4100', anchor='center', text='DMD DDC4100')
                    tree_log_separete.heading('IA Test Pattern Revision', anchor='center',text='IA Test Pattern Revision')
                    tree_log_separete.heading('IA Area Mask Pattern Revision', anchor='center', text='IA Area Mask Pattern Revision')
                    tree_log_separete.heading('IA MAC Address', anchor='center', text='IA MAC Address')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo: 
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.98)
                elif info[4] == "LACS":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=lacs_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no') 
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_Unit Type',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_Unit Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_Boot Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_FPGA Hardware Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_DD Hardware Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_XGA Test Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('LE_Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='Product ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('LE_Unit Type',anchor='center', text='LE_Unit Type')
                    tree_log_separete.heading('LE_Unit Revision', anchor='center', text='LE_Unit Revision')
                    tree_log_separete.heading('LE_Boot Software Version', anchor='center', text='LE_Boot Software Version')
                    tree_log_separete.heading('LE_Software Version', anchor='center', text='LE_Software Version')
                    tree_log_separete.heading('LE_FPGA Hardware Revision', anchor='center', text='LE_FPGA Hardware Revision')
                    tree_log_separete.heading('LE_DD Hardware Revision', anchor='center',text='LE_DD Hardware Revision')
                    tree_log_separete.heading('LE_XGA Test Pattern Revision', anchor='center', text='LE_XGA Test Pattern Revision')
                    tree_log_separete.heading('LE_Area Mask Pattern Revision', anchor='center', text='LE_Area Mask Pattern Revision')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo: 
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.9) 
                elif info[4] == "PE" or info[4] == "PE-Ver2":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=pe_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no') 
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('PE_Boot SoftVersion',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('PE-CPUCoreFPGA0_hard Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('IE_hard Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('PE_MC Control SoftVersion', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('RE-VS_hard Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('PE-FPGA_hard Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('PE_test pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='Product ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('PE_Boot SoftVersion',anchor='center', text='PE_Boot SoftVersion')
                    tree_log_separete.heading('PE-CPUCoreFPGA0_hard Revision', anchor='center', text='PE-CPUCoreFPGA0_hard Revision')
                    tree_log_separete.heading('IE_MC Control SoftVersion', anchor='center', text='IE_MC Control SoftVersion')
                    tree_log_separete.heading('IE_hard Revision', anchor='center', text='IE_hard Revision')
                    tree_log_separete.heading('PE_MC Control SoftVersion', anchor='center', text='PE_MC Control SoftVersion')
                    tree_log_separete.heading('RE-VS_hard Revision', anchor='center',text='RE-VS_hard Revision')
                    tree_log_separete.heading('PE-FPGA_hard Revision', anchor='center', text='PE-FPGA_hard Revision')
                    tree_log_separete.heading('PE_test pattern Revision', anchor='center', text='PE_test pattern Revision')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo: 
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40,borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.98)   
                elif info[4] == "MATE" or info[4] == "MATE3":
                    tree_log_common = ttk.Treeview(logframe1, columns=common_column)
                    tree_log_separete = ttk.Treeview(logframe1, columns=mate_column_log)
                    #列の設定
                    tree_log_common.column('#0', width=0, stretch='no')
                    tree_log_common.column('ID', anchor='center', width=150, stretch='no')
                    tree_log_common.column('Model', anchor='center', width=150, stretch='no')
                    tree_log_common.column('HeadNum', anchor='center', width=100, stretch='no')
                    tree_log_common.column('Direction', anchor='center', width=100, stretch='no')
                    tree_log_separete.column('#0', width=0, stretch='no') 
                    tree_log_separete.column('Machine_No',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Unit Type',anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE FPGA Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE DMD Board Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE App Software Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Interface SoC Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Plot FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Head FPGA Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Unit Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD Type', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('DMD DLPC910 Version', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Test Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE Area Mask Pattern Revision', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('MATE MAC Address', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updater', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Updatedate', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('Customer', anchor='center', width=200, stretch='no')
                    tree_log_separete.column('blank', anchor='center', width=100, stretch='no')
                    #列の見出し
                    tree_log_common.heading('#0', text='')
                    tree_log_common.heading('ID', anchor='center', text='Product ID')
                    tree_log_common.heading('Model', anchor='center', text='Model')
                    tree_log_common.heading('HeadNum', anchor='center', text='HeadNum',command=lambda: treeview_sort_column(tree_log_common, tree_log_separete, 'HeadNum', 'Direction'))
                    tree_log_common.heading('Direction', anchor='center', text='Direction',command=lambda: treeview_sort_dircolumn(tree_log_common, tree_log_separete, 'Direction'))
                    tree_log_separete.heading('#0', text='')
                    tree_log_separete.heading('Machine_No',anchor='center', text='Machine_No')
                    tree_log_separete.heading('MATE Unit Type',anchor='center', text='MATE Unit Type')
                    tree_log_separete.heading('MATE FPGA Board Revision', anchor='center', text='MATE FPGA Board Revision')
                    tree_log_separete.heading('MATE DMD Board Revision', anchor='center', text='MATE DMD Board Revision')
                    tree_log_separete.heading('MATE App Software Version', anchor='center', text='MATE App Software Version')
                    tree_log_separete.heading('MATE Interface SoC Version', anchor='center', text='MATE Interface SoC Version')
                    tree_log_separete.heading('MATE Plot FPGA Version', anchor='center',text='MATE Plot FPGA Version')
                    tree_log_separete.heading('MATE Head FPGA Version', anchor='center', text='MATE Head FPGA Version')
                    tree_log_separete.heading('DMD Unit Type', anchor='center', text='DMD Unit Type')
                    tree_log_separete.heading('DMD Type', anchor='center', text='DMD Type')
                    tree_log_separete.heading('DMD DLPC910 Version', anchor='center', text='DMD DLPC910 Version')
                    tree_log_separete.heading('MATE Test Pattern Revision', anchor='center', text='MATE Test Pattern Revision')
                    tree_log_separete.heading('MATE Area Mask Pattern Revision', anchor='center', text='MATE Area Mask Pattern Revision')
                    tree_log_separete.heading('MATE MAC Address', anchor='center', text='MATE MAC Address')
                    tree_log_separete.heading('Updater', anchor='center', text='Updater')
                    tree_log_separete.heading('Updatedate', anchor='center', text='Updatedate',command=lambda: treeview_sort_timecolumn(tree_log_common, tree_log_separete, 'Updatedate'))
                    tree_log_separete.heading('Customer', anchor='center', text='Customer')
                    tree_log_separete.heading('blank', anchor='center', text='')
                    #レコードの追加
                    count =0
                    for head in headsinfo: 
                        tree_log_common.insert(parent='', index='end', iid=count, values=(info[0], info[1], head[1], head[2]))
                        tree_log_separete.insert(parent='', index='end', iid=count, values=(info[3], head[3], head[4], head[5], head[6], head[7], head[8], head[9], head[10], head[11], head[12], head[13], head[14], head[15], head[16], head[17], info[2], ""))
                        count +=1
                    # Styleの設定
                    style = ttk.Style()
                    # Treeviewの選択時の背景色をデフォルトと同じにする
                    style.map('Treeview', 
                            background=[('selected', style.lookup('Treeview', 'background'))],
                            foreground=[('selected', style.lookup('Treeview', 'foreground'))])
                    style.configure('Treeview', font=(None, 15), rowheight=40, borderwidth=0)
                    style.configure("Treeview.Heading", font=(None, 12))
                    # Treeviewの枠線を非表示にする
                    style.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
                    #スクロールバーの追加
                    logscrollbarx = ttk.Scrollbar(logframe1, orient=HORIZONTAL)
                    tree_log_separete.configure(xscrollcommand = logscrollbarx.set)
                    logscrollbarx[ 'command' ] = tree_log_separete.xview
                    logscrollbarx.place(relx=0.52, rely=0.98, relwidth=0.466)
                    #マウスホイールの同期
                    tree_log_common.bind("<MouseWheel>", on_mouse_wheel)
                    tree_log_separete.bind("<MouseWheel>", on_mouse_wheel)
                    # ウィジェットの配置
                    tree_log_common.place(relheight=1.0,relwidth=0.52)
                    tree_log_separete.place(relx=0.52, relheight=1.0, relwidth=0.48)
                    # 縦スクロールバー
                    logscrollbary = ttk.Scrollbar(logframe1, orient=VERTICAL)
                    tree_log_separete.configure(yscroll = logscrollbary.set)
                    logscrollbary[ 'command' ] = sync_tree_log_separete_yview
                    logscrollbary.place(relx=0.985, relheight=0.98)   
                else:
                    tmsg.showerror("Error",f"機種の登録データはありません{machine_id}.") 
        logsave_button['state'] = 'normal'
    else:
        tmsg.showerror("Error","製番が空白です")

 #列に合わせての入れかえ(ID)      
def treeview_sort_idcolumn(tree_common, tree_separate, id):
    global sort_id
    if sort_id:    
        l = [(tree_common.set(item, id), item) for item in tree_common.get_children('')]
        #ID順に並べ替え
        l.sort(key=lambda t: int(t[0]))
        sorted_data = []
        for _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_id = False
    else:
        l = [(tree_common.set(item, id), item) for item in tree_common.get_children('')]
        #Head番号順に並べ替え
        l.sort(key=lambda t: int(t[0]),reverse=True)
        sorted_data = []
        for _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_id = True
        
 #列に合わせての入れかえ(Customer)      
def treeview_sort_customercolumn(tree_common, tree_separate, customer):
    global sort_customer
    if sort_customer:    
        l = [(tree_separate.set(item, customer), item) for item in tree_separate.get_children('')]
        #Customer順に並べ替え
        l.sort(key=lambda t: t[0],reverse=True)
        sorted_data = []
        for _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_customer = False
    else:
        l = [(tree_separate.set(item, customer), item) for item in tree_separate.get_children('')]
        #Customer順に並べ替え
        l.sort(key=lambda t: t[0],reverse=False)
        sorted_data = []
        for _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_customer = True
        
#列に合わせての入れかえ(ヘッド番号)      
def treeview_sort_column(tree_common, tree_separate, headnum, direction):
    global sort_ascending
    if sort_ascending:    
        l = [(tree_common.set(item, headnum), tree_common.set(item, direction), item) for item in tree_common.get_children('')]
        #Head番号順に並べ替え
        l.sort(key=lambda t: t[1])
        l.sort(key=lambda t: int(t[0]))
        sorted_data = []
        for _, _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_ascending = False
    else:
        l = [(tree_common.set(item, headnum), tree_common.set(item, direction), item) for item in tree_common.get_children('')]
        #Head番号順に並べ替え
        l.sort(key=lambda t: t[1])
        l.sort(key=lambda t: int(t[0]),reverse=True)
        sorted_data = []
        for _, _, item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_ascending = True
 
#列に合わせての入れかえ(Direction)      
def treeview_sort_dircolumn(tree_common, tree_separate, direction):
    global sort_direction
    if sort_direction:    
        l = [(tree_common.set(item, direction), item) for item in tree_common.get_children('')]
        #Direction番号順に並べ替え
        l = sorted(l, key=lambda x: (x[0] == 'Left'))
        sorted_data = []
        for _,  item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_direction = False
    else:
        l = [(tree_common.set(item, direction), item) for item in tree_common.get_children('')]
        #Direction順に並べ替え
        l = sorted(l, key=lambda x: (x[0] != 'Left'))
        sorted_data = []
        for _,  item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_direction = True
        
#列に合わせての入れかえ(time)      
def treeview_sort_timecolumn(tree_common, tree_separate, time):
    global sort_time
    if sort_time:    
        l = [(tree_separate.set(item, time), item) for item in tree_separate.get_children('')]
        # 日時でソートする
        l = sorted(l, key=lambda x: datetime.strptime(x[0], '%Y/%m/%d %H:%M'))

        sorted_data = []
        for _,  item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_time = False
    else:
        l = [(tree_separate.set(item, time), item) for item in tree_separate.get_children('')]
        #Direction番号順に並べ替え
        l = sorted(l, key=lambda x: datetime.strptime(x[0], '%Y/%m/%d %H:%M'),reverse=True)
        sorted_data = []
        for _,  item in l:
            item_data_common = tree_common.item(item)["values"]
            item_data_separate = tree_separate.item(item)["values"]
            sorted_data.append((item_data_common, item_data_separate))

        for i, (data_common, data_separate) in enumerate(sorted_data):
            tree_common.item(tree_common.get_children()[i], values=data_common)
            tree_separate.item(tree_separate.get_children()[i], values=data_separate)
        sort_time = True
        
#IDlistの並び替え関数
def treeview_main_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # reverse sort next time
    tv.heading(col, text=col, command=lambda _col=col: \
                    treeview_main_sort_column(tv, _col, not reverse))

#選択行のIDを取得    
def select_record(event):
    # 選択行の判別
    record_id = tree_common.focus()
    # 選択行のレコードを取得
    record_values = tree_common.item(record_id, 'values')
    entry_text1_var.set(record_values[0])    

#タイトルの変更
def change_title(mode):
    root.title(root.winfo_toplevel().title() + "_" + mode)
 
 #ユーザーのログイン   
def cehck_userinfo(conn,username,passward):
    global authority
    global loginname
    cur=conn.cursor()
    # Head情報の存在チェック
    cur.execute('''SELECT * FROM Users WHERE id = ? AND passward = ? ''', (username, passward))
    info= cur.fetchall()
    if info:
        change_title("Admin")
        authority = True
        loginname = info[0][2]
        [mainframe1.tkraise(), mainframe2.tkraise()]
    else:
        authority = False
        tmsg.showerror("Error","ユーザー名かパスワードが間違っています")
        
#ログインパスワード無し        
def no_check_userinfo():
    global authority
    authority = False
    [mainframe1.tkraise(), mainframe2.tkraise()]
    
#パスワードの伏字切り替え    
def pass_display():
    if pass_entry['show'] == '*':
        pass_entry['show'] = ''
    else:
        pass_entry['show'] = '*'
        
def user_reg(conn):
    for i in range(len(userlist)):
        insert_users(conn, userlist[i][0], userlist[i][1])

#ディレクトリをzip形式で保存        
def dir_to_zip(path, name):
    try:
        dt_now = datetime.now()
        # zipの保存場所（save_dir/name の階層）
        save_dir = os.path.join("save_folder", name)
        # 保存フォルダが存在しない場合、作成する
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        # zipファイル名（現在の日時を使ってユニークな名前を生成）
        zip_filename = dt_now.strftime('%Y%m%d_%H%M%S') + '.zip'
        # ファイルの一時置き場
        temp_dir = 'temp_copy_folder'
        # フォルダの内容を一時的な場所にコピー
        shutil.copytree(path, temp_dir)
        # コピーしたフォルダをzip形式で圧縮
        shutil.make_archive(os.path.join(save_dir, zip_filename.replace('.zip', '')), 'zip', temp_dir)
        # 一時的に作成したフォルダを削除
        shutil.rmtree(temp_dir)
    except:
        tmsg.showerror("Error","ディレクトリをZip形式で保存するのに失敗しました")
    
#ファイルをzip形式で保存
def file_to_zip(path, name):
    try:
        dt_now = datetime.now()
        # zipの保存場所（save_dir/name の階層）
        save_dir = os.path.join("save_folder", name)
        # 保存フォルダが存在しない場合、作成する
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        # zipファイル名（現在の日時を使ってユニークな名前を生成）
        zip_filename = dt_now.strftime('%Y%m%d_%H%M%S') + '.zip'
        zip_path = os.path.join(save_dir, zip_filename)
        # ファイルをZip形式で圧縮
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(path, os.path.basename(path))
    except:
        tmsg.showerror("Error","ファイルをZip形式で保存するのに失敗しました")
        
        
#Customerlistの作成
def reroad_customerlist(cur):
    global customerlist
    try:
        cur.execute("SELECT customer From Machines")
        info = cur.fetchall()
        for data in info:
            if data[0] in customerlist:
                continue
            else:
                customerlist.append(data[0])
    except:
        tmsg.show("Error","顧客リストの取得に失敗しました。")
        
#Customerlistの作成
def reroad_customerlistall(cur):
    global customerlistall
    try:
        cur.execute("SELECT customer From Machines")
        info = cur.fetchall()
        for data in info:
            if data[0] in customerlistall:
                continue
            else:
                customerlistall.append(data[0])
        cur.execute("SELECT customer From BasicInfo")
        info2 = cur.fetchall()
        for data in info2:
            if data[0] in customerlistall:
                continue
            else:
                customerlistall.append(data[0])
    except:
        tmsg.show("Error","顧客リストの取得に失敗しました。")

#テキストの削除        
def clear_textbox():
    ddtextbox.delete('1.0', END)
        

#-----------------カスタムダイアログ一覧-----------------#      
# カスタムダイアログの作成
class SideSelectionDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("機種選択")
        self.side_var = StringVar(value="Left")
        self.geometry("400x250")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.result = ""
        
        Label(self, text="装置方向 (Left or Right)", font=('', 20)).pack(pady=10)
        radiostyle.configure('infoside.TRadiobutton', font=('', 20))  # フォントを設定
        ttk.Radiobutton(self, text="Left", variable=self.side_var, value="Left", style='infoside.TRadiobutton').pack(anchor=W)
        ttk.Radiobutton(self, text="Right", variable=self.side_var, value="Right", style='infoside.TRadiobutton').pack(anchor=W)
        Button(self, text="OK", font=('Helvetica',20), command=self.on_ok).pack(pady=10)
         
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def on_ok(self):
        self.result = self.side_var.get()
        self.destroy()

#カスタムダイアログの召喚
def show_side_selection_dialog(parent):
    dialog = SideSelectionDialog(parent)
    parent.wait_window(dialog)
    return dialog.result

# カスタムダイアログエンジンの作成
class SideEngineSelectionDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("エンジン情報の選択")
        self.dire_var = StringVar(value="Both")
        self.format_var = StringVar(value="All")
        self.geometry("500x400")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.dire =""
        self.format =""
        
        selectframe1 = LabelFrame(self, text="装置方向 (Both or Left or Right)", foreground="green", font=('', 18))
        selectframe1.place(x=10, y=10)
        radiostyle.configure('engineinfo.TRadiobutton', font=('Helvetica', 20))  # フォントを設定
        ttk.Radiobutton(selectframe1, text="Both", variable=self.dire_var, value="Both", style='engineinfo.TRadiobutton').pack(anchor=W)
        ttk.Radiobutton(selectframe1, text="Left", variable=self.dire_var, value="Left", style='engineinfo.TRadiobutton').pack(anchor=W)
        ttk.Radiobutton(selectframe1, text="Right", variable=self.dire_var, value="Right", style='engineinfo.TRadiobutton').pack(anchor=W)
        selectframe2 = LabelFrame(self,text="表示形式 選択 (All or Head0 )", foreground="green", font=('', 18))
        selectframe2.place(x=10,y=180)
        ttk.Radiobutton(selectframe2, text="All", variable=self.format_var, value="All", style='engineinfo.TRadiobutton').pack(anchor=W)
        ttk.Radiobutton(selectframe2, text="Head0", variable=self.format_var, value="Head0", style='engineinfo.TRadiobutton').pack(anchor=W)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=230,y=320)
    
    #画面の中心を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def on_ok(self):
        self.dire = self.dire_var.get()
        self.format = self.format_var.get()
        self.destroy()

#カスタムダイアログエンジンの召喚
def show_engine_selection_dialog(parent):
    dialog = SideEngineSelectionDialog(parent)
    parent.wait_window(dialog)
    return dialog.dire, dialog.format

# カスタムダイアログ顧客情報の作成(ALL)
class CustomerSelectionDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("顧客情報")
        self.geometry("360x360")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.customer = StringVar()
        self.machinetype = StringVar()
        self.modelnum = StringVar()
        self.customer = ""
        self.machinetype = ""
        self.modelnum = ""
        #顧客入力
        self.infomation_label = Label(self, text="顧客名を入力してください", font=('',18))
        self.customer_text = ttk.Combobox(self,  justify="center", width = 26, font=('',18), values=sorted(customerlistall))
        self.infomation_label.place(x=10, y=10)
        self.customer_text.place(x=10, y=60)
        #号機入力
        self.modelnum_label = Label(self, text="号機を入力してください", font=('',18))
        self.modelnum_text = Entry(self,font=('',18), justify="center", width=27)
        self.modelnum_label.place(x=10, y=105)
        self.modelnum_text.place(x=10, y=155)
        #機種label生成
        self.infomation_model = Label(self ,text='機種を選択してください', font=('',18))
        self.infomation_model.place(x=10,y=200)
        #コンボボックスの生成
        self.customer_combobox = ttk.Combobox(self,  justify="center", width = 26, font=('',18), values=machinelist,state="readonly")
        self.customer_combobox.set(machinelist[0])
        self.customer_combobox.place(x=10,y=250)
        okbutton = Button(self, text="OK", font=('Helvetica',20), command=self.on_ok)
        okbutton.place(x=145,y=300)
        
        #号機入力制限の設定
        vcmd = (self.register(self.restrictmodelnum), '%P')
        self.modelnum_text.config(validate="key", validatecommand=vcmd)
    
    #画面の中心座標を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.customer = self.customer_text.get()
        self.machinetype = self.customer_combobox.get()
        self.modelnum = self.modelnum_text.get()
        self.destroy()
        
    #号機入力制限
    def restrictmodelnum(self, string):
        alnumReg = re.compile(r'^[0-9\-\(\)]+$')
        if alnumReg.match(string) or string == "":
            return True
        else:
            return False

#カスタムダイアログ顧客情報の召喚(ALL)
def show_customer_selection_dialog(parent):
    dialog = CustomerSelectionDialog(parent)
    parent.wait_window(dialog)
    return dialog.customer, dialog.machinetype , dialog.modelnum

# カスタムダイアログ顧客と機種
class CustomerandMachineTypeDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Customer Information")
        self.geometry("360x260")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.customer = StringVar()
        self.machinetype = StringVar()
        self.customer = ""
        self.machinetype = ""
        #顧客入力
        self.infomation_label = Label(self, text="顧客名を入力してください", font=('', 18))
        self.customer_text = ttk.Combobox(self,  justify="center", width = 26, font=('',18), values=sorted(customerlistall))
        self.infomation_label.place(x=10, y=10)
        self.customer_text.place(x=10, y=60)

        #機種label生成
        self.infomation_model = Label(self,text='機種を選択してください', font=('', 18))
        self.infomation_model.place(x=10, y=105)
        #コンボボックスの生成
        self.customer_combobox = ttk.Combobox(self,  justify="center", width = 27, font=('', 18), values=machinelist,state="readonly")
        self.customer_combobox.set(machinelist[0])
        self.customer_combobox.place(x=10, y=150)
        okbutton = Button(self, text="OK", font=('Helvetica', 20) , command=self.on_ok)
        okbutton.place(x=145, y=200)
    
    #画面の中心座標を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.customer = self.customer_text.get()
        self.machinetype = self.customer_combobox.get()
        self.destroy()
        
#カスタムダイアログ顧客と機種の召喚
def show_customerandmachinetype_dialog(parent):
    dialog = CustomerandMachineTypeDialog(parent)
    parent.wait_window(dialog)
    return dialog.customer, dialog.machinetype

# カスタムダイアログ顧客情報と号機の作成
class CustomerandModelnumDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("顧客情報")
        self.geometry("365x260")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.customer = StringVar()
        self.modelnum = StringVar()
        self.customer = ""
        self.modelnum = ""
        #顧客入力
        self.infomation_label = Label(self, text="顧客名を入力してください", font=('', 18))
        self.customer_text = ttk.Combobox(self,  justify="center", width = 26, font=('',18), values=sorted(customerlistall))
        self.infomation_label.place(x=10, y=10)
        self.customer_text.place(x=10, y=60)
        #号機入力
        self.modelnum_label = Label(self, text="号機を入力してください", font=('', 18))
        self.modelnum_text = Entry(self,font=('', 18), justify="center", width=28)
        self.modelnum_label.place(x=10, y=105)
        self.modelnum_text.place(x=10, y=155)
        okbutton = Button(self, text="OK", font=('', 20), command=self.on_ok)
        okbutton.place(x=145, y=200)
        
        #号機入力制限の設定
        vcmd = (self.register(self.restrictmodelnum), '%P')
        self.modelnum_text.config(validate="key", validatecommand=vcmd)
    
    #画面の中心座標を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.customer = self.customer_text.get()
        self.modelnum = self.modelnum_text.get()
        self.destroy()
        
    #号機入力制限
    def restrictmodelnum(self, string):
        alnumReg = re.compile(r'^[0-9\-\(\)]+$')
        if alnumReg.match(string) or string == "":
            return True
        else:
            return False

#カスタムダイアログ顧客情報と号機の召喚
def show_customerandmodel_dialog(parent):
    dialog = CustomerandModelnumDialog(parent)
    parent.wait_window(dialog)
    return dialog.customer, dialog.modelnum

# カスタムダイアログ号機と機種
class ModelandMachineTypeDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("顧客情報")
        self.geometry("360x260")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.machinetype = StringVar()
        self.modelnum = StringVar()
        self.machinetype = ""
        self.modelnum = ""
        #号機入力
        self.modelnum_label = Label(self, text="号機を入力してください", font=('', 18))
        self.modelnum_text = Entry(self,font=('', 18), justify="center", width=25)
        self.modelnum_label.place(x=10, y=10)
        self.modelnum_text.place(x=10, y=60)
        
        #号機入力制限の設定
        vcmd = (self.register(self.restrictmodelnum), '%P')
        self.modelnum_text.config(validate="key", validatecommand=vcmd)
        
        #機種label生成
        self.infomation_model = Label(self, text='機種を選択してください', font=('', 18))
        self.infomation_model.place(x=10, y=105)
        #コンボボックスの生成
        self.customer_combobox = ttk.Combobox(self, justify="center", width = 24, font=('', 18), values=machinelist,state="readonly")
        self.customer_combobox.set(machinelist[0])
        self.customer_combobox.place(x=10, y=150)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=145,y=200)
    
    #画面の中心座標を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.machinetype = self.customer_combobox.get()
        self.modelnum = self.modelnum_text.get()
        self.destroy()
        
    #号機入力制限
    def restrictmodelnum(self, string):
        alnumReg = re.compile(r'^[0-9\-\(\)]+$')
        if alnumReg.match(string) or string == "":
            return True
        else:
            return False

#カスタムダイアログ号機と機種
def show_modelandmachinetype_dialog(parent):
    dialog = ModelandMachineTypeDialog(parent)
    parent.wait_window(dialog)
    return dialog.machinetype , dialog.modelnum


# カスタムダイアログ機種選択の作成
class MachineTypeSelectionDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("機種選択")
        self.geometry("280x160")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.machinetype = StringVar()
        self.machinetype = ""
        #機種label生成
        self.infomation_model = Label(self,text='機種を選択してください', font=('', 18))
        self.infomation_model.place(x=10, y=10)
        #コンボボックスの生成
        self.customer_combobox = ttk.Combobox(self,  justify="center", width = 19, font=('', 18), values=machinelist,state="readonly")
        self.customer_combobox.set(machinelist[0])
        self.customer_combobox.place(x=10, y=55)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=110,y=100)
    
    #画面の中心を割り出す処理 
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
       
    def on_ok(self):
        self.machinetype = self.customer_combobox.get()
        self.destroy()

#カスタムダイアログ機種選択の召喚
def show_machinetype_selection_dialog(parent):
    dialog = MachineTypeSelectionDialog(parent)
    parent.wait_window(dialog)
    return  dialog.machinetype

# カスタムダイアログ顧客のみの作成
class CustomerOnlyDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("顧客情報")
        self.geometry("320x160")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)   
        self.customer = StringVar()
        self.customer = ""
        self.infomation_label = Label(self, text="顧客を入力してください", font=('', 18))
        self.customer_text = ttk.Combobox(self,  justify="center", width = 23, font=('',18), values=sorted(customerlistall))
        self.infomation_label.place(x=10, y=10)
        self.customer_text.place(x=10, y=50)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=120, y=100)
    
    #画面の中心座標を割り出す処理
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.customer = self.customer_text.get()
        self.destroy()

#カスタムダイアログ顧客のみの召喚
def show_customeronly_selection_dialog(parent):
    dialog = CustomerOnlyDialog(parent)
    parent.wait_window(dialog)
    return dialog.customer

# カスタムダイアログ号機のみの作成
class ModelnumOnlyDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("顧客情報")
        self.geometry("370x160")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)   
        self.modelnum = StringVar()
        self.modelnum = ""
        self.infomation_label = Label(self, text="号機を入力してください", font=('', 18))
        self.modelnum_text = Entry(self,font=('', 18), justify="center", width=28)
        self.infomation_label.place(x=10, y=10)
        self.modelnum_text.place(x=10, y=50)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=150, y=100)
        #号機入力制限の設定
        vcmd = (self.register(self.restrictmodelnum), '%P')
        self.modelnum_text.config(validate="key", validatecommand=vcmd)
    
    #画面の中心座標を割り出す処理
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.modelnum = self.modelnum_text.get()
        self.destroy()
        
    #号機入力制限
    def restrictmodelnum(self, string):
        alnumReg = re.compile(r'^[0-9\-\%\(\)]+$')
        if alnumReg.match(string) or string == "":
            return True
        else:
            return False

#カスタムダイアログ号機のみの召喚
def show_modelnumonly_dialog(parent):
    dialog = ModelnumOnlyDialog(parent)
    parent.wait_window(dialog)
    return dialog.modelnum

# カスタムダイアログ顧客とエンジンタイプ
class CustomerandEngineTypeDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("検索条件")
        self.geometry("340x260")
        self.resizable(False,False)    # リサイズ不可に設定
        self.bind("<Visibility>", self.centering_main_window)
        self.customer = StringVar()
        self.enginetype = StringVar()
        self.customer = ""
        self.enginetype = ""
        #顧客入力
        self.infomation_label = Label(self, text="顧客を入力してください", font=('',18))
        self.customer_text = ttk.Combobox(self,  justify="center", width = 24, font=('', 18), values=sorted(customerlist))
        self.infomation_label.place(x=10, y=10)
        self.customer_text.place(x=10, y=60)

        #機種label生成
        self.infomation_model = Label(self,text='エンジンを選択してください', font=('', 18),)
        self.infomation_model.place(x=10, y=105)
        #コンボボックスの生成
        self.engine_combobox = ttk.Combobox(self,  justify="center", width = 24, font=('', 18), values=enginelist,state="readonly")
        self.engine_combobox.set(enginelist[0])
        self.engine_combobox.place(x=10, y=150)
        okbutton = Button(self, text="OK", font=('Helvetica', 20), command=self.on_ok)
        okbutton.place(x=125, y=200)
    
    #画面の中心座標を割り出す処理    
    def centering_main_window(self, event):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
    def on_ok(self):
        self.customer = self.customer_text.get()
        self.enginetype = self.engine_combobox.get()
        self.destroy()
        
#カスタムダイアログ顧客と機種の召喚
def show_customerandenginetype_dialog(parent):
    dialog = CustomerandEngineTypeDialog(parent)
    parent.wait_window(dialog)
    return dialog.customer, dialog.enginetype

#-----------------更新時間取得用のクエリ-----------------#        
query_laiza = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_laiza
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

query_livera = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_livera
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

query_ilia = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_ilia
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

query_lacs = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_lacs
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

query_pe = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_pe
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

query_mate = '''
WITH latest_updates AS (
    SELECT updatedate, loginname
    FROM Heads_mate
    WHERE latestflag = 'Latest' AND headnum = ? AND machines_id = ?   
    UNION ALL  
    SELECT updatedate, loginname
    FROM DPC
    WHERE latestflag = 'Latest' AND machines_id = ?
    UNION ALL
    SELECT updatedate, loginname
    FROM DXPCON
    WHERE latestflag = 'Latest' AND machines_id = ?
)
SELECT loginname, updatedate
FROM latest_updates
ORDER BY updatedate DESC
LIMIT 1;
'''

#-----------------顧客情報取得用のクエリ-----------------# 
query_customerinfo =  'SELECT * FROM Machines Where customer LIKE ? AND enginetype = ? ORDER BY customer ASC'
        
#データベース名
dbname = 'Di_Version_Control.db'
# dbname = '\\\\192.168.7.33\\cad\\各部門\\Di装置制御部\\データベース\\Di_Version_Control.db'
tree_common = None
tree_separate = None
scrollbarx = None
scrollbary = None
tree_log_common = None
tree_log_separete = None
logscrollbarx = None
logscrollbary = None
conn = sqlite3.connect(dbname)
create_main_table(conn)
create_subtable_laiza(conn)
create_subtable_livera(conn)
create_subtable_ilia(conn)
create_subtable_lacs(conn)
create_subtable_pe(conn)
create_subtable_mate(conn)
create_subtable_dpc(conn)
create_subtable_dxpcon(conn)
create_subtable_excellist(conn)
create_subtable_basicinfo(conn)
create_subtable_user(conn)
sort_id = False
sort_customer = True
sort_ascending = False
sort_direction = False
sort_time = True
id_flag = False
authority = False
loginname = ""

#-----------------treeviewのColumnに使用-----------------#
#固定列
common_column = ('ID', 'Model', 'HeadNum','Direction')

#laiza用の列
laiza_column = ('Machine_No', 'DXPCON Version', 'DPC Version', 'LI Unit Type', 'LI Vector Process Board Revision', 'LI Head Process Board Revision', 'LI App Software Version',
                'LI Vector Process FPGA Version', 'LI Intersection Process FPGA Version', 'LI Head Process FPGA Version', 'LI FHD Test Pattern Version', 'LI XGA Test Pattern Version', 
                'LI Area Mask Pattern Version', 'LI MAC Address', 'DMD Unit Type', 'DMD Type', 'DMD DDC4100 Version', 'Updater', 'Updatedate','Customer','blank')

#livera用の列
livera_column =('Machine_No', 'DXPCON Version', 'DPC Version', 'LV Unit Type', 'LV Board Revision', 'LV Boot Software Version', 'LV App Software Version', 'LV FPGA Version',
                'LV FHD Test Pattern Version', 'LV XGA Test Pattern Version', 'LV Area Mask Pattern Version', 'LV MAC Address', 'DMD Unit Type', 'DMD Board Revision', 'DMD FPGA Version', 
                'DMD Type', 'DMD DDC4100 Version', 'Updater', 'Updatedate','Customer','blank')
#ilia用の列
ilia_column = ('Machine_No', 'DXPCON Version', 'DPC Version','IA Unit Type', 'IA USB Interface Board Revision', 'IA Head Process Board Revision', 'IA App Software Version',
                'IA Interface SoC Version', 'IA Plot FPGA Revision', 'DMD Unit Type', 'DMD Type', 'DMD DDC4100', 'IA Test Pattern Revision', 'IA Area Mask Pattern Revision', 'IA MAC Address',
                 'Updater', 'Updatedate', 'Customer', 'blank')

#lacs用の列
lacs_column = ('Machine_No', 'DXPCON Version', 'DPC Version', 'LE_Unit Type', 'LE_Unit Revision', 'LE_Boot Software Version', 'LE_Software Version', 'LE_FPGA Hardware Revision',
                   'LE_DD Hardware Revision', 'LE_XGA Test Pattern Revision', 'LE_Area Mask Pattern Revision', 'Updater', 'Updatedate', 'Customer', 'blank')

#PE用の列
pe_column = ('Machine_No', 'DXPCON Version', 'DPC Version', 'PE_Boot SoftVersion','PE-CPUCoreFPGA_hard Revision','IE_MC Control SoftVersion', 'IE_hard Revision',
             'PE_MC Control SoftVersion', 'RE-VS_hard Revision', 'PE-FPGA_hard Revision', 'PE_test pattern Revision', 'Updater', 'Updatedate', 'Customer', 'blank')

#MATE用の列
mate_column = ('Machine_No', 'DXPCON Version', 'DPC Version', 'MATE Unit Type', 'MATE FPGA Board Revision', 'MATE DMD Board Revision', 'MATE App Software Version', 'MATE Interface SoC Version', 
               'MATE Plot FPGA Version', 'MATE Head FPGA Version', 'DMD Unit Type', 'DMD Type', 'DMD DLPC910 Version', 'MATE Test Pattern Revision', 'MATE Area Mask Pattern Revision', 'MATE MAC Address',
                'Updater', 'Updatedate', 'Customer', 'blank')

#main用の列
main_column = ('ID', 'Model', 'Customer', 'Machine_No', 'Engine', 'blank')

#固定列
common_column_log = ('ID', 'Model')

#laiza用の列
laiza_column_log = ('Machine_No', 'LI Unit Type', 'LI Vector Process Board Revision', 'LI Head Process Board Revision', 'LI App Software Version','LI Vector Process FPGA Version',
                    'LI Intersection Process FPGA Version', 'LI Head Process FPGA Version', 'LI FHD Test Pattern Version', 'LI XGA Test Pattern Version', 
                    'LI Area Mask Pattern Version', 'LI MAC Address', 'DMD Unit Type', 'DMD Type', 'DMD DDC4100 Version', 'Updater', 'Updatedate', 'Customer', 'blank')

#livera用の列
livera_column_log =('Machine_No', 'LV Unit Type', 'LV Board Revision', 'LV Boot Software Version', 'LV App Software Version', 'LV FPGA Version','LV FHD Test Pattern Version', 
                    'LV XGA Test Pattern Version', 'LV Area Mask Pattern Version', 'LV MAC Address', 'DMD Unit Type', 'DMD Board Revision', 'DMD FPGA Version', 
                    'DMD Type', 'DMD DDC4100 Version', 'Updater', 'Updatedate', 'Customer', 'blank')

#ilia用の列
ilia_column_log = ('Machine_No', 'IA Unit Type', 'IA USB Interface Board Revision', 'IA Head Process Board Revision', 'IA App Software Version','IA Interface SoC Version', 
                   'IA Plot FPGA Revision', 'DMD Unit Type', 'DMD Type', 'DMD DDC4100', 'IA Test Pattern Revision', 'IA Area Mask Pattern Revision', 'IA MAC Address',
                    'Updater','Updatedate','Customer','blank')

#lacs用の列
lacs_column_log = ('Machine_No', 'LE_Unit Type', 'LE_Unit Revision', 'LE_Boot Software Version', 'LE_Software Version', 'LE_FPGA Hardware Revision',
                   'LE_DD Hardware Revision', 'LE_XGA Test Pattern Revision', 'LE_Area Mask Pattern Revision', 'Updater', 'Updatedate', 'Customer', 'blank')

#PE用の列
pe_column_log = ('Machine_No', 'PE_Boot SoftVersion','PE-CPUCoreFPGA0_hard Revision','IE_MC Control SoftVersion', 'IE_hard Revision',
                   'PE_MC Control SoftVersion', 'RE-VS_hard Revision', 'PE-FPGA_hard Revision', 'PE_test pattern Revision', 'Updater', 'Updatedate', 'Customer', 'blank')

#MATE用の列
mate_column_log = ('Machine_No', 'MATE Unit Type', 'MATE FPGA Board Revision', 'MATE DMD Board Revision', 'MATE App Software Version', 'MATE Interface SoC Version',
                   'MATE Plot FPGA Version', 'MATE Head FPGA Version', 'DMD Unit Type', 'DMD Type', 'DMD DLPC910 Version', 'MATE Test Pattern Revision', 'MATE Area Mask Pattern Revision',
                   'MATE MAC Address', 'Updater', 'Updatedate', 'Customer', 'blank')

#dpc用の列
DPC_colimn_log = ('Machine_No', 'DPC Version', 'Updater', 'Updatedate', 'Customer', 'blank')

#dxpcon用の列
DXPCON_colmn_log = ('Machine_No', 'DXPCON Version', 'Updater', 'Updatedate', 'Customer', 'blank')

#-----------------ユーザーリスト-----------------#
userlist = [["hiroyuki-tanaka", "田中 広之"], ["t-ono", "尾野 太紀"], ["c-yamagishi", "山岸 智洋"], ["n-okada", "岡田 典之"], ["f-okada", "岡田 文雄"], ["h-morita", "森田 英行"],
            ["a-kikushima", "菊島 晶博"], ["e-nishida", "西田 栄治"], ["s-arai", "新井 郷史"]]

#-----------------コンボボックスのリストに使用-----------------#
laizalist = []
liveralist = []
ilialist = []
lacslist = []
pelist =[]
pe2list = []
matelist = []
mate3list = []
machinelist =[]
enginelist = []
customerlist = []
customerlistall = []

#ユーザーリストの登録
user_reg(conn)

#カレントディレクトリへのPathを取得
current = os.getcwd()
path = os.path.join(current, 'Machinetype.txt')
try:
    append_list(path)
except:
    tmsg.showerror("Error","カレントディレクトリにMachinetype.txtがありません")
    sys.exit()

#-----------CSのエクセルファイル確認処理----------#
directory_path = "\\\\192.168.54.62\\d\\DICS部\\Di装置一覧"
readflag = True
timestamp = ""
if os.path.isdir(directory_path):
    result = file_get(directory_path)
    regname = os.path.basename(result)
    timestamp =datetime.fromtimestamp(os.path.getmtime(result))
    timestamp = timestamp.strftime('%Y/%m/%d %H:%M')
    readflag = inquiry(conn,timestamp)
    
if readflag != True:
    try:       
        df = pd.read_excel(result, sheet_name='DI装置一覧', usecols=[1,6,7,9,11,22])
        #補正を実行
        df = fill_merged_cells(df)
        delete_basicinfo(conn)
        create_subtable_basicinfo(conn)
        insert_excellist(conn, regname, timestamp)
        count = 0
        for row in df.itertuples():
            if count >=1:
                if row[1] != "廃棄":
                    insert_basicinfo(conn, row[4], row[3], row[2], row[5],row[6])
            count += 1
    except :
        tmsg.showinfo("Attention ","Excelファイル使用中の為,読み込みができませんでした")

root = TkinterDnD.Tk()
#-----------Window作成------------#
root.title("Di Version Control Database" +"_"+Version) # 画面タイトル設定
root.geometry('1370x880')  # 画面サイズ設定
root.minsize(1370,880)
#root.resizable(False,True)    # リサイズ不可に設定

#----------Frameを定義------------#
# mainframe1 = Frame(root, width=960, height=880, borderwidth=2, relief='solid')
# mainframe2 = Frame(root, width=410, height=880, borderwidth=2, relief='solid')
# logframe1  = Frame(root, width=960, height=880, borderwidth=2, relief='solid')
# logframe2  = Frame(root, width=410, height=880, borderwidth=2, relief='solid')
# loginframe1 = Frame(root, width=960, height=880, borderwidth=0, relief='solid')
# loginframe2  = Frame(root, width=410, height=880, borderwidth=0, relief='solid')

mainframe1 = Frame(root, borderwidth=1, relief='solid')
mainframe2 = Frame(root,  borderwidth=1, relief='solid')
logframe1  = Frame(root,  borderwidth=1, relief='solid')
logframe2  = Frame(root,  borderwidth=1, relief='solid')
loginframe1 = Frame(root,  borderwidth=0, relief='solid')
loginframe2  = Frame(root,  borderwidth=0, relief='solid')

#Frameサイズを固定
# mainframe1.propagate(False)
# mainframe2.propagate(False)
# logframe1.propagate(False)
# logframe2.propagate(False)
# loginframe1.propagate(False)
# loginframe2.propagate(False)

#Frameを配置
# mainframe1.grid(row=0, column=0)
# mainframe2.grid(row=0, column=1)
# logframe1.grid(row=0, column=0)
# logframe2.grid(row=0, column=1)
# loginframe1.grid(row=0, column=0)
# loginframe2.grid(row=0, column=1)

mainframe1.place(relheight=1.0,relwidth=0.7)
mainframe2.place(relx=0.7, relheight=1.0,relwidth=0.3)
logframe1.place(relheight=1.0,relwidth=0.7)
logframe2.place(relx=0.7,relheight=1.0,relwidth=0.3)
loginframe1.place(relheight=1.0,relwidth=0.7)
loginframe2.place(relx=0.7,relheight=1.0,relwidth=0.3)

#----------Widgetの配置------------#
#製番label生成
label_la = Label(mainframe2, text='製番', font=('', 18))
label_la.place(x=25, y=0)

#製番入力TextBox生成
entry_text1_var = StringVar()
entry_text1_var.trace_add("write", on_entry_text1_change)
vc = root.register(at_least_charlog)
entry_text1 = Entry(mainframe2, font=('',18), justify="center", width=15, textvariable=entry_text1_var, validate="key", validatecommand=(vc, "%P"))
entry_text1.place(x=25, y=28)

#ファイルラベルの生成
label_file = Label(mainframe2, text='ファイル・フォルダ', font=('', 18))
label_file.place(x=25, y=58)

#drag＆dropできるテキストボックスの生成
ddtextbox = Text(mainframe2,font=('', 18), width=25, height=5)
ddtextbox.drop_target_register(DND_FILES)
ddtextbox.dnd_bind("<<Drop>>", get_textpath)
ddtextbox.place(x=25, y=88)

#クリアボタンの生成
clear_button = Button(mainframe2, width = 4, height=2, font=('', 16), text='クリア', command=clear_textbox)
clear_button.place(x=338, y = 120)

#検索ボタンの生成
search_button = Button(mainframe2, width = 15, height=2, font=('', 18), text='検索', state='disabled', command=lambda:display_machine_info(conn,entry_text1.get()))
search_button.place(x=25, y=220)

#検索要素の選択
searchframe = LabelFrame(mainframe2, text="検索タイプ", foreground="green", font=('', 12))
searchframe.place(x=270, y=213)
#ラジオボタンの生成
searchsettei = IntVar()
searchsettei.set(0)
# スタイルの設定左よせ
radiostyle = ttk.Style()
radiostyle.configure('info.TRadiobutton', font=('', 14))  # フォントを設定
search0 = ttk.Radiobutton(searchframe, value=0, text='製番', variable=searchsettei, style='info.TRadiobutton', command=lambda:on_id())
search0.pack(anchor="w")
search1 = ttk.Radiobutton(searchframe,value=1, text='顧客名', variable=searchsettei, style='info.TRadiobutton', command=lambda:on_cus())
search1.pack(anchor="w")
search2 = ttk.Radiobutton(searchframe,value=2, text='号機番号', variable=searchsettei, style='info.TRadiobutton', command=lambda:on_cus())
search2.pack(anchor="w")

#登録ボタンの生成
register_button = Button(mainframe2, width = 15, height=2, font=('',18), text='登録', state='disabled', command=lambda:Click_reg_data(conn,ddtextbox.get(0.,END).rstrip('\n'), entry_text1.get()))
register_button.place(x=25, y=300)
#更新ボタンの生成
update_botton = Button(mainframe2, width = 15, height=2, font=('',18), text='更新', state='disabled', command=lambda:Click_update_data(conn,ddtextbox.get(0.,END).rstrip('\n'), entry_text1.get()))
update_botton.place(x=25, y=385)

#更新要素の選択
listupdateframe = LabelFrame(mainframe2, text="更新タイプ", foreground="green", font=('', 12))
listupdateframe.place(x=270,y=355)
customer_button_check = BooleanVar()
customercheck = Checkbutton(listupdateframe, text="顧客名    ", font=('', 12) , variable=customer_button_check)
customer_button_check.set(False)
customercheck.pack(anchor="w")
machinetype_button_check = BooleanVar()
machinetype_check = Checkbutton(listupdateframe, text="機種", font=('', 12) , variable=machinetype_button_check)
machinetype_button_check.set(False)
machinetype_check.pack(anchor="w")
modelnum_button_check = BooleanVar()
modelnum_check = Checkbutton(listupdateframe, text="号機", font=('', 12) , variable=modelnum_button_check)
modelnum_button_check.set(False)
modelnum_check.pack(anchor="w")

#一覧ボタンの生成
listView_botton = Button(mainframe2, width = 15, height=2, font=('', 18),text='機種 一覧', command=lambda:list_view(conn))
listView_botton.place(x=25, y=470)
#フレームの生成
listframe = LabelFrame(mainframe2, text="表示タイプ 設定", foreground="green", font=('', 12))
listframe.place(x=270, y=463)
#ラジオボタンの生成
settei = IntVar()
settei.set(0)
radio_0=ttk.Radiobutton(listframe, value=0, text='All', variable=settei, style='info.TRadiobutton')
radio_0.pack(anchor="w")
radio_1=ttk.Radiobutton(listframe, value=1, text='Head0', variable=settei, style='info.TRadiobutton')
radio_1.pack(anchor="w")
#削除ボタンの生成
delete_button = Button(mainframe2, width = 15, height=2, font=('',18), text='削除', state='disabled', command=lambda:delete_data(conn,entry_text1.get()))
delete_button.place(x=25, y=555)
#表示要素の選択
listelementframe = LabelFrame(mainframe2, text="表示要素", foreground="green", font=('', 12))
listelementframe.place(x=270, y=548)
machine_no_button_check = BooleanVar()
Machine_No_check = Checkbutton(listelementframe, text = "号機", font=('', 12) , variable=machine_no_button_check,command=reload)
machine_no_button_check.set(True)
Machine_No_check.pack(anchor="w")
dxpcon_button_check = BooleanVar()
Dxpconcheck = Checkbutton(listelementframe, text="DXPCON ", font=('', 12) , variable=dxpcon_button_check,command=reload)
dxpcon_button_check.set(True)
Dxpconcheck.pack(anchor="w")
dpc_button_check = BooleanVar()
Dpccheck = Checkbutton(listelementframe, text="DPC",font=('', 12) , variable=dpc_button_check,command=reload)
dpc_button_check.set(True)
Dpccheck.pack(anchor="w")
engine_button_check = BooleanVar()
Enginecheck = Checkbutton(listelementframe, text="ENGINE",font=('', 12) , variable=engine_button_check,command=reload)
engine_button_check.set(True)
Enginecheck.pack(anchor="w")

#ID一覧ボタンの生成
idlist_button = Button(mainframe2, width = 15, height=2, font=('', 18), text='ID 一覧', command=lambda:display_idlist(conn))
idlist_button.place(x=25,y=640)
#更新履歴の表示ボタンの生成
history_button = Button(mainframe2, width = 15, height=2, font=('', 18), text='履歴', command=lambda:[logframe1.tkraise(), logframe2.tkraise()])
history_button.place(x=25,y=725)
#csv出力ボタンの生成
save_button = Button(mainframe2, width = 15, height=2, font=('', 18), text='CSV 出力',state='disabled', command=lambda:save_csv())
save_button.place(x=25,y=805)

#--------logframe内のウィジット記述---------#
#製番label生成
ID_label = Label(logframe2, text='製番', font=('', 18))
ID_label.place(x=25, y=0)

#製番入力TextBox生成
entry_textlog_var = StringVar()
vc = root.register(at_least_char)
entry_textlog = Entry(logframe2, font = ('', 18), justify="center", textvariable=entry_textlog_var, width=15, validate="key", validatecommand=(vc, "%P"))
entry_textlog.place(x=25, y=35)

#フレームの生成
loglistframe = LabelFrame(logframe2, text="表示要素 選択", foreground="green", font=('', 16))
loglistframe.place(x=25, y=75)
#ラジオボタンの生成
logsettei = IntVar()
logsettei.set(0)
# スタイルの設定左よせ
logradiostyle = ttk.Style()
logradiostyle.configure('infolog.TRadiobutton', font=('', 16)) #フォントを設定
logradio_0 = ttk.Radiobutton(loglistframe, value=0, text='DXPCON', variable=logsettei, style='infolog.TRadiobutton')
logradio_0.pack(anchor="w")
logradio_1 = ttk.Radiobutton(loglistframe, value=1, text='DPC', variable=logsettei, style='infolog.TRadiobutton')
logradio_1.pack(anchor="w")
logradio_2 = ttk.Radiobutton(loglistframe, value=2, text='ENGINE', variable=logsettei, style='infolog.TRadiobutton')
logradio_2.pack(anchor="w")

#検索ボタンの生成
logsearch_button = Button(logframe2, width = 15, height=2, font=('',18), text='検索', state='disabled', command=lambda:displaylog_machine_info(conn,entry_textlog.get()))
logsearch_button.place(x=25, y=215)

#Mainに戻るボタンの生成
history_button = Button(logframe2, width = 15, height=2, font=('',18), text='メインに戻る', command=lambda:[mainframe1.tkraise(), mainframe2.tkraise()])
history_button.place(x=25, y=310)

#csv出力ボタンの生成
logsave_button = Button(logframe2, width = 15, height=2, font=('',18), text='CSV 出力', state='disabled', command=lambda:logsave_csv())
logsave_button.place(x=25, y=405)

#--------loginframe内のウィジット記述---------#
itemframe = LabelFrame(loginframe1,text = "ログイン", foreground="green", font=('', 20))
itemframe.place(x=550, y=200)
name_label = Label(itemframe, text = "ユーザー名", font =('', 18))
name_label.pack(anchor="w", pady=10, padx=5)
name_entry = Entry(itemframe, font = ('',18), justify = "center",width=25)
#name_entry.insert(0,"a-kikushima")
name_entry.pack(anchor="w", padx=5)
pass_label = Label(itemframe, text= "パスワード", font =('', 18))
pass_label.pack(anchor="w", pady=10, padx=5)
pass_entry = Entry(itemframe, font = ('',18), justify = "center",width=25)
#pass_entry.insert(0,"hirai285")
pass_entry.pack(anchor="w", padx=5)
pass_entry['show'] = '*'
passcheck = Checkbutton(itemframe, text="パスワードを表示する",font=('', 12), command=lambda:pass_display())
passcheck.pack(anchor="w", padx=5)
login_button = Button(itemframe, width = 15, height=2, font=('',18), text='ログイン', command=lambda:cehck_userinfo(conn, name_entry.get(), pass_entry.get()))
login_button.pack(anchor="center", pady=20)
no_login_button = Button(itemframe, width = 15, height=2, font=('',18), text='ゲストログイン',command=lambda:no_check_userinfo())
no_login_button.pack(anchor="center", pady=10)

loginframe1.tkraise()
loginframe2.tkraise()
#ウィンドウの表示#
root.mainloop()