import tkinter as tk
import xlwings as xw
import datetime
from tkinter import ttk
from tkinter.messagebox import showinfo
import pandas as pd

ws= tk.Tk()
ws.title("wokschedule")
#=========================視窗設定========================================
ws.geometry("400x400+700+300")  #大小+位置 
#ws.resizable(False,False)        #不可更改大小   
ws.config(bg = "#323232")        #顏色 
ws.attributes("-topmost",True)   #置頂

#============================功能設定========================================
#click_count = 0   button計數器
hoursinf = [[],[],[],[],[],[],[],[],[],[],[],[],[]]
hours_binf = [[],[],[],[],[],[],[],[],[],[],[],[],[]]
titleinf = []
title_binf = []
nameinf = []
name_binf = []

#==========================寫入EXCEL=========================================
def output():
#    global click_count
#    click_count += 1;
    app = xw.App(visible=True,add_book=False)  #開excel
    wb = app.books.open('班表範例.xlsx')        #讀檔
    
    wb.sheets['使用說明'].range('C2').value = yearen.get()
    wb.sheets['使用說明'].range('E2').value = monthen.get()
    wb.sheets['使用說明'].range('H2').value = shopen.get()
    
    #上半月工作表寫入
    gethoursdict = {}
    for c in range(15):
        gethourslist = []
        for r in range(13):
            gethourslist.append(hoursinf[r][c].get())        #同一行整理成一個list
        gethoursdict.setdefault(str(c+1),gethourslist)            #整理成一個字典
    gethoursdf = pd.DataFrame(gethoursdict)                       #整理成一個DataFrame
    wb.sheets['1~15'].range('C3').options(pd.DataFrame,header = False,index = False,expend = 'table').value = gethoursdf
    
    getnamelist = []
    gettitlelist = []
    for r in range(13):
        getnamelist.append(nameinf[r].get())                #將name的資料存成list
        gettitlelist.append(titleinf[r].get())              #將title的資料存成list
    wb.sheets['1~15'].range('A3').options(transpose = True).value = getnamelist
    wb.sheets['1~15'].range('B3').options(transpose = True).value = gettitlelist
    
    #下半月工作表寫入
    gethours_bdict = {}
    for c in range(16):
        gethours_blist = []
        for r in range(13):
            gethours_blist.append(hours_binf[r][c].get())        #同一行整理成一個list
        gethours_bdict.setdefault(str(c+16),gethours_blist)            #整理成一個字典
    gethours_bdf = pd.DataFrame(gethours_bdict)                       #整理成一個DataFrame
    wb.sheets['16~31'].range('C3').options(pd.DataFrame,header = False,index = False,expend = 'table').value = gethours_bdf 
    
    getname_blist = []
    gettitle_blist = []
    for r in range(13):
        getname_blist.append(name_binf[r].get())                #將name的資料存成list
        gettitle_blist.append(title_binf[r].get())              #將title的資料存成list
    wb.sheets['16~31'].range('A3').options(transpose = True).value = getname_blist
    wb.sheets['16~31'].range('B3').options(transpose = True).value = gettitle_blist
    
    
    wb.save(r'C:\Users\chan\Desktop\python code\a\%s班表%s%s.xlsx' %(shopen.get(),yearen.get(),monthen.get()))  #另存新檔(若已存在則覆蓋之)
    app.quit()
#=========================勞基法檢查========================================
def check():
#==============一天最多十小時=================   
    massage = []
    for r in range(13):
        for c in range(15):
            try :  
                text= hoursinf[r][c].get().split('/')
                hours = 0                   # 將/兩邊的字串分開後再做相加運算
                for s in range(len(text)):
                    hourscheck = eval(text[s]) 
                    if  hourscheck >= 0:
                        hours = hours + (12-hourscheck)
                    else:
                        hours += abs(hourscheck)
            except: 
                print("", end = "")
            if hours > 10 :
                massage.append(nameinf[r].get()+'的第'+str(c+1)+'日超過十小時'+'\n')  #將所有提示組成字串
    for r in range(13):
        for c in range(16):
            try :  
                text= hours_binf[r][c].get().split('/')
                hours = 0                   # 將/兩邊的字串分開後再做相加運算
                for s in range(len(text)):
                    hourscheck = eval(text[s]) 
                    if  hourscheck >= 0:
                        hours = hours + (12-hourscheck)
                    else:
                        hours += abs(hourscheck)
            except: 
                print("", end = "")
            if hours > 10 :
                massage.append(name_binf[r].get()+'的第'+str(c+16)+'日超過十小時'+'\n')
    
#=================一周不得超過48小時不得超過七天(最多六天)===================================
    getnamelist = []
    for r in range(13):
        getnamelist.append(nameinf[r].get())                #將上半月name的資料存成list           
    gethoursdict = {}
    for c in range(15):
        gethourslist = []
        for r in range(13):
            gethourslist.append(hoursinf[r][c].get())             #同一行整理成一個list
        gethoursdict.setdefault(str(c+1),gethourslist)            #整理成一個字典
    gethoursdf = pd.DataFrame(gethoursdict,index = getnamelist) #將上半月整理成一個DataFrame 列標題為人名 行標題為1~15        

    getname_blist = []
    for r in range(13):
        getname_blist.append(name_binf[r].get())                #將下半月name的資料存成list
    gethours_bdict = {}
    for c in range(16):
        gethours_blist = []
        for r in range(13):
            gethours_blist.append(hours_binf[r][c].get())                 #同一行整理成一個list
        gethours_bdict.setdefault(str(c+16),gethours_blist)               #整理成一個字典
    gethours_bdf = pd.DataFrame(gethours_bdict,index = getname_blist)   #將上半月整理成一個DataFrame 列標題為人名 行標題為16~31
    
    dfmerge = gethoursdf.join(gethours_bdf,how = 'outer').fillna('')          #將上下半月兩個DataFrame以列為參考合併成一個大的DataFrame 並將預設的NAN改成空白''
    
    for r in (x for x in dfmerge.index if x!='') :              #使用 index裡非空白的做成list給 r
        for a in range(25) :
            hours = 0
            count7 = 0
            for c in range(a,a+7):
                try :
                    hourscount = 0                          # 將/兩邊的字串分開後再做相加運算
                    text= dfmerge.loc[r][c].split('/')      # .loc['索引'][n] : 索引那列的第n-1個字串     .iloc[n1][n2] : 第(n1-1)列的第(n2-1)個字串
                    for s in range(len(text)):
                        hourscheck = eval(text[s]) 
                        if  hourscheck >= 0:
                            hours = hours + (12-hourscheck)
                        else:
                            hours += abs(hourscheck)
                        hourscount = hours
                        if hourscount > 0:               #hourscount 算完每天就歸零  hours算完七天才歸零
                            count7 += 1
                except: 
                    print('')
            if hours > 48 :
                massage.append(r+'的第'+str(a+1)+'~'+str(a+7)+'日總和超過四八小時(請計算加班'+str(hours-48)+'小時)\n')
            if count7 == 7 :
                massage.append(r+'的第'+str(a+1)+'~'+str(a+7)+'日違反不得連續上班7日\n')
                
    if len(massage)==0:
        showinfo("勞基法檢查",'均符合勞基法') 
    else:
        showmassage=''
        for r in range(len(massage)):
            showmassage += massage[r]            
        showinfo("勞基法檢查",showmassage)    
    
#=========================介面設定========================================
shop = tk.Label(text = "門市",font = "微軟正黑體 18",bg = "#323232",fg = "white",anchor = "e")        
shopen = tk.Entry(width = 12,font = "微軟正黑體 12",justify = "center")
year = tk.Label(text = "年份",font = "微軟正黑體 18",bg = "#323232",fg = "white",justify = "right")        
yearen = ttk.Combobox(values=["2020","2021","2022","2023","2024","2025","2026","2027","2028","2029","2030",
                              "2031","2032","2033","2034","2035","2036","2037","2038","2039"],
                      width = 4,font = "微軟正黑體 12",justify = "center",state = "readonly")
month = tk.Label(text = "月份",font = "微軟正黑體 18",bg = "#323232",fg = "white",justify = "right")        
monthen =  ttk.Combobox(values=["01","02","03","04","05","06","07","08","09","10","11","12"],width = 4,font = "微軟正黑體 12",justify = "center",state = "readonly")

shop.place(x = 170, y = 50)
shopen.place(x = 140, y = 85)
year.place(x = 170, y = 120)
yearen.place(x = 170, y = 155)
month.place(x = 170, y = 190)
monthen.place(x = 170, y = 225)


def weekday():
    ws.geometry("1300x950+280+30") 
    shop.grid(row = 0, column = 1)
    shopen.grid(row = 0, column = 2, columnspan =2)
    year.grid(row = 0, column = 4)
    yearen.grid(row = 0, column = 5)
    month.grid(row = 0, column = 6)
    monthen.grid(row = 0, column = 7)
    weekbtn.destroy() 
    weekbtn2 = tk.Button(command = weekday, text = "月份更新",width = 11,font = "微軟正黑體 8",relief = "raised")
    weekbtn2.grid(row = 0, column = 8,columnspan = 2)
 
    for r in range(13):
        for c in range(15):
            hoursinf[r].append(tk.Entry(width = 7,font ="微軟正黑體 12",justify = "center"))       
            hoursinf[r][c].grid(row = r+3,column = c+4)
    for r in range(13):
        for c in range(16):
            hours_binf[r].append(tk.Entry(width = 7,font ="微軟正黑體 12",justify = "center"))
            hours_binf[r][c].grid(row = r+20,column = c+4)

    schedule = tk.Label(text = "上半月排班表",font = "微軟正黑體 20",bg = "#323232",fg = "white") 
    schedule.grid(row = 1, column = 1, columnspan =20)

    name = tk.Label(text = "姓名",font = "微軟正黑體 12", width = 8,bg = "#323232",fg = "white")
    name.grid(row = 2, column = 1, columnspan =2)

    title = tk.Label(text = "職別",font = "微軟正黑體 12", width = 8,bg = "#323232",fg = "white")
    title.grid(row = 2, column = 3)

    for r in range(13):
        titleinf.append(ttk.Combobox(values=["工時", "實習組長","組長","儲備幹部","儲備店長","副店長","店長",],width = 7,font ="微軟正黑體 12",justify = "center"))
        titleinf[r].grid(row = r+3, column = 3)

    for r in range(13):
        nameinf.append(tk.Entry(width = 10,font ="微軟正黑體 12", bg = "skyblue",justify = "center"))
        nameinf[r].grid(row = r+3, column = 1,columnspan =2)
      
    schedule_b = tk.Label(text = "下半月排班表",font = "微軟正黑體 20",bg = "#323232",fg = "white") 
    schedule_b.grid(row = 18, column = 1, columnspan =20)

    name_b = tk.Label(text = "姓名",font = "微軟正黑體 12", width = 8,bg = "#323232",fg = "white")
    name_b.grid(row = 19, column = 1, columnspan =2)

    title_b = tk.Label(text = "職別",font = "微軟正黑體 12", width = 8,bg = "#323232",fg = "white")
    title_b.grid(row = 19, column = 3)

    for r in range(13):
        title_binf.append(tk.ttk.Combobox(values=["工時", "實習組長","組長","儲備幹部","儲備店長","副店長","店長",],width = 7,font ="微軟正黑體 12",justify = "center"))
        title_binf[r].grid(row = r+20, column = 3)
    
    for r in range(13):
        name_binf.append(tk.Entry(width = 10,font ="微軟正黑體 12", bg = "skyblue",justify = "center"))
        name_binf[r].grid(row = r+20, column = 1,columnspan =2)

    blankr1 = tk.Label(text = "",font = "微軟正黑體 5",bg = "#323232")
    blankr1.grid(row = 17, column = 4)  

    blankr2 = tk.Label(text = "",font = "微軟正黑體 5",bg = "#323232")
    blankr2.grid(row = 33, column = 4)  

    blankc3 = tk.Label(text = "",bg = "#323232",width = 3)
    blankc3.grid(row = 0, column = 0) 

    download = tk.Button(command = output, text = "建立Excel", width = 15,font = "微軟正黑體 18",relief = "raised")
    download.grid(row = 34, column = 15, columnspan =8)
    
    dict = {1: "一",2:"二",3:"三",4:"四",5:"五",6:"六",7:"日",}
    if len(yearen.get()) != 0 and len(monthen.get()) != 0:
        for d in range(15):
            checkday = dict[datetime.date(int(yearen.get()),int(monthen.get()),d+1).isoweekday()]
            date = tk.Label(text = str(d+1)+ " " + checkday,font ="微軟正黑體 12",width = 6,bg = "#323232",fg = "white")
            date.grid(row = 2, column = d+4)
        for d in range(16):
            try:
                checkday2 = dict[datetime.date(int(yearen.get()),int(monthen.get()),d+16).isoweekday()]      #判斷是否為錯誤日期 EX:2/30，如為錯誤則印空白
                date_b = tk.Label(text = str(d+16)+ " " + checkday2,font ="微軟正黑體 12",width = 6,bg = "#323232",fg = "white")
            except :
                date_b = tk.Label(text = " ",font ="微軟正黑體 12",width = 6,bg = "#323232",fg = "white")
            date_b.grid(row = 19, column = d+4)
            
    ckeckbtn = tk.Button(command = check, text = "勞基法檢查", width = 15,font = "微軟正黑體 18",relief = "raised")
    ckeckbtn.grid(row = 34, column = 1, columnspan =8)
    
weekbtn = tk.Button(command = weekday, text = "建立",width = 5,font = "微軟正黑體 20",relief = "raised")
weekbtn.place(x = 155, y = 285)


ws.mainloop()