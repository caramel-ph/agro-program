import sys, os
from tkinter import *
from tksheet import Sheet
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import sqlite3 as sql
import pandas as pd
from PIL import ImageTk, Image
from PIL import Image, ImageGrab, ImageOps
from tksheet import *

from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil
import re

import csv
from os.path import normpath
import io
import xlsxwriter
import subprocess

import plotly.graph_objects as go
from plotly.subplots import make_subplots

root = Tk()
# Получаем ширину и высоту экрана
scwidth = root.winfo_screenwidth()
scheight = root.winfo_screenheight()
# Закрываем экземпляресли окно не нужно
root.destroy()



sheetindicator1 = 0
sheetindicator2 = 0
tab2indicator1 = 0
canvindicator1 = 0
Years_to_show = []

def round_up(x):
    try:
        return float(ceil(x))
    except:
        return x

def open_file():                ###### Menu Открыть таблицу
    global sheet1
    global sheetindicator2
    filepath = filedialog.askopenfilename()
    messagebox.showinfo("GUI Python", "Загружен Excel-файл : "+ filepath)

    cxn = sql.connect('mydb.db')
    wb = pd.read_excel(filepath)
    wb = wb.replace('None', '')
    column_names_pd = wb.head()

    column_names_pd_file = open("column_names_pd.txt", "w+")
    for r in column_names_pd:
        column_names_pd_file.write(str(r) + " ")
    column_names_pd_file.close()
    subprocess.check_call(["attrib","+H","column_names_pd.txt"])

    
    wb = wb.rename(columns={"Год": "годгод"})
    wb = wb.fillna('')  
    print(wb)
    print(wb.head())
    wb.to_sql(name='mytable',con=cxn,if_exists='replace',index=False)
    cxn.commit()
    crsheet1()
    if sheetindicator2:
        sheet2.destroy()
        sheetindicator2=0
   
def save_click():               ###### Menu  Сохранить
    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]

    mdata = sheet1.get_sheet_data(get_header = False, get_index = False)
    wb = pd.DataFrame(mdata, columns=column_names)
    print(wb)
    wb.to_sql(name='mytable',con=cxn,if_exists='replace',index=False)
    cxn.commit()

    crsheet1()
    messagebox.showinfo("saveing", "Файл успешно сохранен")

def save_click_tab2():
    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]
    
    mdata = sheet1.get_sheet_data(get_header = False, get_index = False)
    wb = pd.DataFrame(mdata, columns=column_names)

    wb.to_sql(name='mytable',con=cxn,if_exists='replace',index=False)
    cxn.commit()

def delms():                #########   "Мастер удаления"
    global selected
    global delt
    global t2
    global combobox

    addw = Toplevel()
    addw.title('Добавить данные')
    addw.geometry('500x300')
    addw.resizable(False, False)

    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]

    t1= Label(addw, text="выберете параметр для удаления")
    t1.place(x=50, y=50)
    t11= Label(addw, text="введите значение удаляемого параметра")
    t11.place(x=50, y=80)

    combobox = ttk.Combobox(addw,values=column_names)
    combobox.place(x=300, y=50)
    selected = combobox.get()
    print(selected)

    t2 = Entry(addw)
    t2.place(x=300, y=80)
    delt = t2.get()
    print(delt)

    addw.del_btn = ttk.Button(addw,text='удалить из базы данных соотв. объекты',command=objdel)
    addw.del_btn.place(x=220, y=120)

    window.update()
    addw.grab_set()
    addw.focus_set()
    addw.mainloop()

def objdel():
    global selected
    global delt

    selected = combobox.get()
    delt = t2.get()
    print("DELETE FROM mytable WHERE " + str(selected) + " = " + str(delt))
    cur.execute("DELETE FROM mytable WHERE " + str(selected) + " = " + "'"+str(delt)+"'")

    messagebox.showinfo('hi', 'удалены все строки, где ' + str(selected) + ' = ' + str(delt))
    cxn.commit()
    crsheet1()

def crsheet1():         #######  Вход!!! Загрузка таблицы
    print("crsheet1")
    global sheet1
    global sheetindicator1
    global art
    
    try:
        canv.destroy()
    except:
        print("")

    if sheetindicator1:
        sheet1.destroy()
    cur.execute("SELECT * FROM mytable")
    art = cur.fetchall()

    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]

    sheet1 = Sheet(frame1, width = int(0.7*scwidth), height=400, headers = column_names,data = [[f"{art[r][c]}" for c in range(len(art[0]))] for r in range(len(art))]) 
    sheet1.enable_bindings()        # редактируется таблица . есть контекстное меню  
    sheet1.pack(side=LEFT)
    sheet1.extra_bindings([("cell_select", cell_select)])

    sheetindicator1=1
    
def cell_select(response):      ###########  ######    Click  on tabl_1
    save_click_tab2()
    #print("cell_select")
    global sheet1
    global sheet2
    global sheetindicator1
    global sheetindicator2
    global canvindicator1
    global canv
    global imeg
    global seldata
    global column_names
    global Years_to_show

    cur.execute("SELECT * FROM mytable")
    art = cur.fetchall()
    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]

    if sheetindicator2:
        sheet2.destroy()

    if canvindicator1:
        canv.destroy()

    sel = response["selected"]

    row = sel.row
    col = sel.column


    cur.execute("SELECT * FROM mytable WHERE rowid = "+str(row+1))
    datarow = cur.fetchall()
    print(str(column_names[col])+" = "+"'"+str(datarow[0][col])+"'")
    cur.execute("SELECT * FROM mytable WHERE "+str(column_names[col])+" = "+"'"+str(datarow[0][col])+"'")
    seldata = cur.fetchall()

    Years_to_show = []
    for i in seldata:
        Years_to_show.append(i[0])

    sheet2 = Sheet(frame2,width = 400,height=400,headers = column_names,data = [[f"{seldata[r][c]}" for c in range(len(seldata[0]))] for r in range(len(seldata))]) 
    
    sheet2.enable_bindings("single_select","right_click_popup_menu")            #таблица2 редактируется

    sheet2.pack(fill=BOTH)

    sheet2.readonly_columns(columns = "all", readonly = True, redraw = False)
    sheet2.popup_menu_del_command(label = None)
    sheet2.popup_menu_add_command("Save sheet", save_sheet_sheet2)
    sheet2.set_all_cell_sizes_to_text()
    sheet2.change_theme("light green")
   

    cur.execute("SELECT фото FROM mytable")

    geometry = window.geometry()
    winwidth = int(geometry.split('x')[0])
    winheight = window.winfo_screenheight()
    try:
        
        img_oo = cur.fetchall()[row]
        print("./photo/"+str(img_oo[0])+".jpg")
        imeg_o = Image.open("./photo/"+str(img_oo[0])+".jpg")
        img = ImageOps.contain(imeg_o, (int(winwidth - 0.7*scwidth), 400))
        imeg = ImageTk.PhotoImage(img)
        canv = Canvas(frame1, width=int(winwidth - 0.7*scwidth), height=400, bg="#ffffff")
        canv.place(x=0,y=0)
        canv.create_image(int((winwidth - 0.7*scwidth)/2), 200, anchor="center", image=imeg)
        canv.pack(anchor=E)    #  anchor=CENTER, expand=1
        canvindicator1=1
    except FileNotFoundError:
        print("Нет фото")
    
    sheetindicator2=1
    pass

def comb():                     ##########   "Объединить с таблицей Excel"
    filepath = filedialog.askopenfilename()
    messagebox.showinfo("GUI Python", "Загружен Excel-файл для объединения : "+ filepath)

    adcxn = sql.connect('mydb.db')
    adwb = pd.read_excel(filepath)
    adwb.to_sql(name='adtable',con=adcxn,if_exists='replace',index=False)
    adcxn.commit()

    cur.execute("SELECT * FROM adtable")
    addata = cur.fetchall()

    cur.execute('PRAGMA table_info("mytable")')
    column_names = [i[1] for i in cur.fetchall()]
    q = '?,'*(len(column_names)-1)+'?'

    for ad in addata:
        cur.execute("INSERT INTO mytable VALUES ("+str(q)+")",(ad))
    messagebox.showinfo('отсчет', 'данные добавлены успешно ')
    cxn.commit()
    crsheet1()

def save_sheet_tab2():              # Menu: 'Экспорт в Excel'
        global tab2
        datacsv = sheet1.get_sheet_data(get_header = True, get_index = False)

        filepathout = filedialog.asksaveasfilename(title = "Сохранить файл",
                                                filetypes = [('xlsx File','.xlsx'),],
                                                defaultextension = ".xlsx",
                                                confirmoverwrite = True)
        
        if not filepathout or not filepathout.lower().endswith((".xlsx", ".tsv")):
            return
        try:
            with xlsxwriter.Workbook(filepathout) as workbook:
                worksheet = workbook.add_worksheet()
                for row_num, data in enumerate(datacsv):
                    worksheet.write_row(row_num, 0, data)
                
        except:
            return

def save_sheet_sheet2():
        global sheet2
        datacsv = sheet2.get_sheet_data(get_header = True, get_index = False)

        filepathout = filedialog.asksaveasfilename(title = "Save sheet as",
                                                filetypes = [('xlsx File','.xlsx'),],
                                                defaultextension = ".xlsx",
                                                confirmoverwrite = True)
        
        if not filepathout or not filepathout.lower().endswith((".xlsx", ".tsv")):
            return
        try:
            with xlsxwriter.Workbook(filepathout) as workbook:
                worksheet = workbook.add_worksheet()
                for row_num, data in enumerate(datacsv):
                    worksheet.write_row(row_num, 0, data)
                
        except:
            return


########## ХОТИМ ПОСТРОИТЬ ГРАФИК ##############
def plot():
    try:
        plotting()
    except:
        filepath = filedialog.askopenfilename()
        messagebox.showinfo("Меню графиков", "Загружен Excel-файл : "+ filepath)

        global pl
        global pt
        pl = pd.read_excel(filepath)
        pt = pd.read_excel(filepath,sheet_name=1,dtype=object)

        plotting()

########## СТРОИМ ГРАФИК ##############
def plotting():
    fig = make_subplots(rows=2, cols=1,specs=[[{"type": "xy","secondary_y": True}],[{"type": "table"}]])

    for year, group in pl.groupby("год"):
        fig.add_trace(go.Scatter(x=group["месяц"], y=group["значение температуры"], name=year,line_shape='spline',
        hovertemplate="%sгод<br>значение температуры=%%{y}<extra></extra>"% year),secondary_y=True,row=1,col=1)

    for year, group in pl.groupby("год"):
        fig.add_trace(go.Bar(x=group["месяц"], y=group["значение осадков"], name=year,
        hovertemplate="%sгод<br>осадки=%%{y}<extra></extra>"% year),secondary_y=False,row=1,col=1)

    fig.for_each_trace(lambda trace: trace.update(visible="legendonly") 
                   if trace.name not in Years_to_show else ())

    fig.add_trace(go.Table(header=dict(values=list(pt.columns),fill_color='grey',align='left',font=dict(color='white', size=12)),cells=dict(values=pt.T,fill=dict(color=['lightgrey', 'white']),align='left'))
    ,row=2,col=1)
    
    # Add figure title
    fig.update_layout(title_text="Температура и осадки")

    # Set x-axis title
    fig.update_xaxes(title_text="месяц")

    # Set y-axes titles
    fig.update_yaxes(title_text="<b>Температура", secondary_y=True)
    fig.update_yaxes(title_text="<b>осадки", secondary_y=False)

    fig.show()


##############  ОКНО ######################

cxn = sql.connect('mydb.db')
cur = cxn.cursor()
cxn.commit()

window = Tk()
window.title("База данных по мискантусу")

# program_directory=sys.path[0]                         # Иконка. Рабочий вариант!!!!!!!!!!!
# window.iconphoto(True, PhotoImage(file=os.path.join(program_directory, "wheat.png")))
try:
    im = Image.open('wheat.png')
    photo = ImageTk.PhotoImage(im)
    window.wm_iconphoto(True, photo)
except Exception:
    0

#window.geometry('1366x768')
window.geometry(f"{scwidth}x{scheight}")
print(scwidth)

frame1 = ttk.Frame(borderwidth=2, relief=RAISED)
frame2 = ttk.Frame(borderwidth=2, relief=RAISED)
frame1 = LabelFrame(text="Коллекция")
frame2 = LabelFrame(text="Выборка")

frame3 = ttk.Frame(borderwidth=2, relief=RAISED)
frame3 = LabelFrame(text="картинка")
frame3.pack(padx=10,pady=10,anchor=NE)

#frame1.pack(padx=10,pady=10,expand=True,anchor=N,fill=X)
frame1.pack(padx=10,pady=10,side=TOP,expand=True,anchor=N,fill=X)
frame2.pack(padx=10,pady=10,expand=True,anchor=S,fill=X)
#-----------------------

window.option_add("*tearOff", FALSE)
main_menu = Menu(window,activeborderwidth=40,background="#4f3b04")
 
main_menu.add_cascade(label="Открыть таблицу", command=open_file)
main_menu.add_cascade(label="Сохранить", command=save_click)
main_menu.add_cascade(label="Объединить с таблицей Excel", command=comb)
main_menu.add_cascade(label='Экспорт в Excel', command=save_sheet_tab2)
main_menu.add_cascade(label="Мастер удаления", command=delms)
main_menu.add_cascade(label="Открыть график погоды", command=plot)
window.config(menu=main_menu)

                        ### первая вкладка
try:
    crsheet1()
except:
    filepath = filedialog.askopenfilename()
    messagebox.showinfo("GUI Python", "Загружен Excel-файл : "+ filepath)

    cxn = sql.connect('mydb.db')
    wb = pd.read_excel(filepath)
    wb.to_sql(name='mytable',con=cxn,if_exists='replace',index=False)
    cxn.commit()
    crsheet1()


window.mainloop()
cxn.close()