from ast import Delete
from email.headerregistry import Group
from email.mime import base
from openpyxl import load_workbook
import openpyxl
import re
import MySQLdb
import pymysql
from config import host,user,password,db_name



try:
    connection = pymysql.connect(
    host=host,
    port=3306,
    user=user,
    password=password,
    database=db_name,
    cursorclass=pymysql.cursors.DictCursor
    )
    print("Connection seccessfully....")
except Exception as ex:
    print("Connection refused....")
    print(ex)


#with connection:
#    with connection.cursor() as cursor:
        # Create a new record
#         sql = "INSERT INTO `teacher` (`idTeacher`,`Name`) VALUES (%s,%s)"
#         val=(10,'Афонин А.Ю.')
#         cursor.execute(sql,val)

    # connection is not autocommit by default. So you must commit to save
    # your changes.
#    connection.commit()
    
con = MySQLdb.connect(host="127.0.0.1",port=3306, user="root", passwd="", db="db4")

cur = connection.cursor()

cur.execute("SELECT Name FROM teacher")

name=[]

for row in cur.fetchall():
    z=[]
    z.append(re.findall(r"'(.*?)'",str(row)))
    try:
        if(z[0][1]):
            name.append(z[0][1])
    except:
        print("Name is not")


    
del z
    

#book=load_workbook(filename="C:/Users/User/Documents/parsing.xlsx")
book=openpyxl.open("C:/Users/User/Documents/parsing.xlsx",read_only=True);

daysz={6:'ПН', 16:'ВТ',26:'СР',36:'ЧТ',46:'ПТ',56:'СБ'}
teacher={'Акифьев И.В.':0,'Гудков А.А.':0,'Дубинин В.Н.':0,'Карамышева Н.С.':0,'Калиниченко Е.И.':0,'Эпп В.В.':0,'Бождай А.С.':0,'Гурин Е.И.':0,'Елфимов А.В.':0,'Кольчугина Е.А.':0,'Афонин А.Ю.':0}
result=''
sheet=book.active
i=17
j=1
checkLast=None # Запоминает последнюю строку в которой был преподователь
disc=0 #Название предмета
Groups=' ' #Номер группы
weeks=0 #Номер недели
days=' ' # День недели
auditorium=''#Номер аудитории
Name=''
try:
    while(sheet[5][i].value!=None):
      # print(sheet[5][i].value)
       Groups=sheet[5][i].value
       while(sheet[5+j][i].value!=sheet[5][i].value):
           if((5+j) in daysz):#Извлечение дня из таблицы
                days=daysz[5+j]
           if(sheet[5+j][i].value):
               for x in name: #Если такое имя есть в таблице 
                    result=re.search(x,sheet[5+j][i].value) 
                    if (result!=None):
                        if(checkLast==None and sheet[5+j-1][i].value!=None):# для первых строк таблицы
                             disc=sheet[5+j-1][i].value
                             weeks=12
                             #auditorium=re.findall(r"(?:\d+-\w+)",sheet[5+j][i].value)
                             auditorium=re.findall(r"[0-9].*[0-9]",sheet[5+j][i].value)
                        elif(checkLast==None and sheet[5+j-1][i].value==None and sheet[5+j][i-1].value!=None):
                            disc=re.findall(r".*?(?=\лб|\лк|$)",sheet[5+j][i].value)
                            weeks=1
                            auditorium=re.findall(r"[0-9].*[0-9]",sheet[5+j][i].value)
                        elif(checkLast==None and sheet[5+j-1][i].value==None and sheet[5+j][i-1].value==None):
                            disc=re.findall(r".*?(?=\лб|\лк|$)",sheet[5+j][i].value)
                            weeks=2
                        elif(checkLast!=(5+j-1) and checkLast!=None and sheet[5+j-1][i].value!=None):# для последующих строк
                             disc=sheet[5+j-1][i].value
                             weeks=12;
                             auditorium=re.findall(r"[0-9].*[0-9]",sheet[5+j][i].value)
                        elif(sheet[5+j-1][i].value==None and checkLast!=(5+j-1) and checkLast!=None):
                             disc=re.findall(r".*?(?=\лб|\лк|$)",sheet[5+j][i].value) 
                             disc=disc[0]
                             auditorium=re.findall(r"[0-9].*[0-9]",sheet[5+j][i].value)
                             if(sheet[5+j][i-1].value!=None):
                                 weeks=1
                             else:
                                 weeks=2
                        elif(checkLast==(5+j-1) and checkLast!=None and sheet[5+j-1][i].value!=None):
                             disc=re.findall(r".*?(?=\лб|\лк|$)",sheet[5+j][i].value) 
                             disc=disc[0]
                             auditorium=re.findall(r"[0-9].*[0-9]",sheet[5+j][i].value)
                             if(sheet[5+j][i-1].value!=None):
                                 weeks=1
                             else:
                                 weeks=2
                        checkLast=5+j 
                        Name=x;
                        print(Groups)
                        print(" ")     
                        print(weeks) 
                        print(" ")
                        print(days) 
                        print(" ")
                        print(Name)
                        print(" ")
                        print(disc)
                        print(" ")
                        print(auditorium)
#                        with connection:
#                            with connection.cursor() as cursor:
                                # Create a new record
#                                sql = "INSERT INTO `Schedule` (`Week`, `Day`,`Class`,`Group`,`Lesson`,`Auditorium`,`Teacher_ID`) VALUES (%s,%s,%s,%s,%s,%s,%s)"
#                                val=(weeks,days,'0',Groups,disc,auditorium,1)
#                                cursor.execute(sql,val)

                            # connection is not autocommit by default. So you must commit to save
                            # your changes.
#                            connection.commit()
                        result=None
                        break
               #print(sheet[5+j][i].value)
           j+=1
       i+=3
       checkLast=None
       j=1
except:
    print("over")
    
for key in teacher:
    print(teacher[key])



with connection:
    with connection.cursor() as cursor:
        # Create a new record
        sql = "INSERT INTO `Schedule` (`Week`, `Day`,`Class`,`Group`,`Lesson`,`Auditorium`,`Teacher_ID`) VALUES (%s,%s,%s,%s,%s,%s,%s)"
        val=(weeks,days,'0',Groups,disc,auditorium,1)
        cursor.execute(sql,val)

        # connection is not autocommit by default. So you must commit to save
        # your changes.
    connection.commit()

#with connection:
#    with connection.cursor() as cursor:
        # Create a new record
#        sql = ("INSERT INTO `Schedule` (`Week`,`Teacher_ID`) VALUES ({},{})".format(weeks,0))
#        val=(1,weeks)
#        cursor.execute(sql)

    # connection is not autocommit by default. So you must commit to save
    # your changes.
#    connection.commit()

con.close()

#disc==''
#if(checkLast!=(5+j-1) and (5+j-1)!=None):
#    disc=sheet[5+j-1][i].value
    
#elif(5+j-1)==None):