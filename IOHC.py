#	IOHC.py 
#   Copyright (C) 2011  Varun Rana<varunrana.in>
#	Developer :Varun Rana, Mayank Saini, Paras Vij
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>
#


import Tkinter
import tkMessageBox
from Tkinter import *
import sqlite3
#from win32com.client import *
import os,sys

mysql_path = "Student.db"
if not os.path.exists(mysql_path):
    connection = sqlite3.connect('Student.db')
    try:
        cursor = connection.cursor()
        sql = "CREATE TABLE `course` (`course_id` INTEGER PRIMARY KEY AUTOINCREMENT , `course_name` VARCHAR(128), `total_seats` INTGER , `allocated_seats` INTGER );"
        cursor.execute(sql)
        cursor.execute("CREATE TABLE `form` (`form_id` INTEGER PRIMARY KEY AUTOINCREMENT, `student_id` VARCHAR(20),`student_name` VARCHAR(128), `priority1` VARCHAR(32), `priority2` VARCHAR(32) NULL, `priority3` VARCHAR(32) NULL, `priority4` VARCHAR(32) NULL )")
        cursor.execute("CREATE TABLE `allocate` (`student_id` VARCHAR(20) PRIMARY KEY,`student_name` VARCHAR(128),`course_id` INTGER)")
        cursor.execute("INSERT INTO `course` (`course_name`,`total_seats`,`allocated_seats`) VALUES ('.NET technologies','40','0')")
        cursor.execute("INSERT INTO `course` (`course_name`,`total_seats`,`allocated_seats`) VALUES ('Linux Administration','40','0')")
        cursor.execute("INSERT INTO `course` (`course_name`,`total_seats`,`allocated_seats`) VALUES ('Web Development with J2EE','40','0')")
        cursor.execute("INSERT INTO `course` (`course_name`,`total_seats`,`allocated_seats`) VALUES ('SQL, PL-SQL, Oracle Architecture','40','0')")
        connection.commit()
    except:
        connection.rollback()    

def SubmitForm():
    StudentId = student_id_var.get()
    StudentId = student_id_var.get()
    StudentName = student_name_var.get()
    FirstPriority = p1var.get()
    SecondPriority = p2var.get()
    ThirdPriority = p3var.get()
    FourthPriority = p4var.get()
    
    
    if(SecondPriority == '--Select--'):
        SecondPriority = 'NULL'
        ThirdPriority = 'NULL'
        FourthPriority = 'NULL'
    if(ThirdPriority == '--Select--'):
        ThirdPriority = 'NULL'
        FourthPriority = 'NULL'
    if(FourthPriority == '--Select--'):
        FourthPriority = 'NULL'
        
        
    if(StudentId == ''):
        tkMessageBox.showinfo("Watch Out!!","Id cannot be left blank")
    elif(StudentName == ''):
        tkMessageBox.showinfo("Watch Out!!","Name cannot be left blank")
    elif(FirstPriority == '--Select--'):
        tkMessageBox.showinfo("Watch Out!!","Select alteast one priority")
    else:
        connection = sqlite3.connect('Student.db')
        try:
            cursor = connection.cursor()
            sql = ("SELECT `student_id` FROM `form` WHERE `student_id`='%s'")%StudentId
            try:
                cursor.execute(sql)
                result = cursor.fetchone()
                if(result[0]==StudentId):    
                    tkMessageBox.showinfo("Watch Out!!","You are already registered")
            except:
                sql = ("INSERT INTO `form`(`student_id`,`student_name`,`priority1`,`priority2`,`priority3`,`priority4`) VALUES('%s','%s','%s','%s','%s','%s')")%(StudentId,StudentName,FirstPriority,SecondPriority,ThirdPriority,FourthPriority)        
                try:
                    cursor.execute(sql)
                    connection.commit()
                    tkMessageBox.showinfo("Yeah", "Form entered Successfully")
                except:
                    tkMessageBox.showinfo("Dude","Something went wrong")
        except:
            print "test"
            connection.rollback()
        
            
def ResetForm():
    student_id_var.set('')
    student_name_var.set('')
    
def CalculateForm():
    StudentId = student_id.get()
    StudentName = student_name.get()
    
"""def xls_read():
    connection = sqlite3.connect('Student.db')
    c=connection.cursor
    XLS_FILE = os.getcwd() + "test.xls"
    ROW_SPAN = (14,21)
    COL_SPAN = (2,7)
    app = Dispatch("Excel.Application")
    app.Visible = True
    ws = app.Workbooks.Open(XLS_FILE).Sheets(1)
    exceldata = [[ws.Cells(row, col).value
        for col in xrange(COL_SPAN[0], COL_SPAN[1])] 
        for row in xrange(ROW_SPAN[0], ROW_SPAN[1])]
    
    try :
    
        for row in exceldata:
            c.execute('INSERT INTO form VALUES (?,?,?,?,?,?)', row)
            print row
        connection.commit()
    except :
        print "hhe"
        connection.rollback
      """  
root = Tk()
root.wm_title("Chitkara IHOC app")

#entry student id
cnv = Canvas(root, width=0, height=10)
cnv.pack()
frame = Frame(root)
frame.pack()
student_id = Label(frame, text="      STUDENT ID :")
student_id.pack( side = LEFT )
cnv = Canvas(root, width=0, height=10)
cnv.pack()
student_id_var = StringVar()
student_id_frame = Entry(frame,textvariable=student_id_var , bd=5)
student_id_frame.pack( side = RIGHT )
student_id_var.set('')

#student name
frame3 = Frame(root)
frame3.pack()
student_name = Label(frame3, text="STUDENT NAME :")
student_name.pack( side = LEFT )
cnv = Canvas(root, width=0, height=10)
cnv.pack()
student_name_var = StringVar()
student_name_frame = Entry(frame3, textvariable = student_name_var, bd = 5)
student_name_frame.pack( side = RIGHT )
student_name_var.set('')
#student course selection
cnv = Canvas(root, width=0, height=5)
cnv.pack()
frame4 = Frame(root)
frame4.pack()
priority1 = Label(frame4, text="PRIORITY1 :")
priority1.pack( side= LEFT )
tup = ("--Select--",".NET technologies","Linux Administration","Web Development with J2EE","SQL, PL-SQL, Oracle Architecture")
p1var = StringVar()
w = Spinbox(frame4, from_=0, to=4, width=25, values = tup, textvariable=p1var)
w.pack( side = RIGHT )

#student course selection2
cnv = Canvas(root, width=0, height=5)
cnv.pack()
frame5 = Frame(root)
frame5.pack()
priority2 = Label(frame5, text="PRIORITY2 :")
priority2.pack( side= LEFT )
tup = ("--Select--",".NET technologies","Linux Administration","Web Development with J2EE","SQL, PL-SQL, Oracle Architecture")
p2var = StringVar()
q = Spinbox(frame5, from_=0, to=4, width=25, values = tup, textvariable = p2var)
q.pack( side = RIGHT)

#student course selection3
cnv = Canvas(root, width=0, height=5)
cnv.pack()
frame6 = Frame(root)
frame6.pack()
priority3 = Label(frame6, text="PRIORITY3 :")
priority3.pack( side= LEFT )
tup = ("--Select--",".NET technologies","Linux Administration","Web Development with J2EE","SQL, PL-SQL, Oracle Architecture")
p3var = StringVar()
e = Spinbox(frame6, from_=0, to=4, width=25, values = tup, textvariable = p3var)
e.pack( side = RIGHT )

#student course selection4
cnv = Canvas(root, width=0, height=5)
cnv.pack()
frame7 = Frame(root)
frame7.pack()
priority4 = Label(frame7, text="PRIORITY4 :")
priority4.pack( side= LEFT )
tup = ("--Select--",".NET technologies","Linux Administration","Web Development with J2EE","SQL, PL-SQL, Oracle Architecture")
p4var = StringVar()
r = Spinbox(frame7, from_=0, to=4, width=25, values = tup, textvariable = p4var)
r.pack( side = RIGHT )
cnv = Canvas(root, width=0, height=5)
cnv.pack()
frame8 = Frame(root)
frame8.pack()
B = Tkinter.Button(frame8, text ="SUBMIT", command = SubmitForm)
B.pack( side = LEFT )
cnv = Canvas(root, width=5, height=0)
cnv.pack()
b = Tkinter.Button(frame8, text ="RESET", command = ResetForm)
b.pack( side = RIGHT )
cnv = Canvas(root, width=5, height=0)
cnv.pack()
v = Tkinter.Button(frame8, text ="CALCULATE", command = CalculateForm)
v.pack()
cnv = Canvas(root, width=5, height=0)
cnv.pack()
root.mainloop()

