'''
After a workbook is connected to OPERAD and required tables are selected, columns are displayed as dimension/measure.
This code helps in importing the column comments using the data dictionary. Before running the code, make sure to have
the workbook connected to some datasoure and the tables are selected. The workbook should be saved in some location on your
device. Then run the code on that workbook using your OPERAD credentials and reopen to see the changes in effect. After 
reopening, if you hover over any of the dimension/measure you should be able to see the associated comments provided the
comments exist for that column in the DB.

Version : 1.0 
'''

#import necessary libraries
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import *
from tkinter import ttk
from tkinter.ttk import *
from ttkthemes import themed_tk as themes
from PIL import ImageTk, Image
import tkinter.messagebox
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement, Comment, tostring
import cx_Oracle
import sql_metadata

def browse():          #function for browsing and locating the workbook
    filename = fd.askopenfilename()
    e.insert(tk.END,filename)
    
def waitingMsg():
    if(L3["text"]==""):
        L3["text"]="Working on it...Please wait"
    elif(L3["text"]=="Working on it...Please wait"):
        L3["text"]=""
    root.update()

def myClick():         #function for importing column descriptions to the workbook from OPERAD
    if not e1.get() or not e2.get() or not e.get() or not hostnameEntry.get() or not portEntry.get() or not serviceEntry.get():
        tkinter.messagebox.showinfo("","Please fill all the fields")
        return
        
    custom=0            #custom flag(0 means the workbook doesn't use custom SQL query to connect to datasource and 1 means it uses the custom SQL query)
    
    tree = ET.parse(e.get())                        #XML element tree parser
    root = tree.getroot()
    datasources = root.find('datasources')
    datasource = datasources.find('datasource')
    
    for i in datasource.iter('relation'):           #sets custom flag to 0 or 1
        if(i.get('name')=='Custom SQL Query'):
            custom=1
            break
        else:
            custom=0
    
    if(custom==0):              
        tree = ET.parse(e.get())
        root = tree.getroot()
        datasources = root.find('datasources')
        datasource = datasources.find('datasource')
        
        columns = []            #stores column tags from workbook XML
        tableName_sec=[]        #stores table names other than the first connected table in the datasource 
        primary=[]              #stores the table name of the first connected table
        match_col=[]            #stores the column names which match among all the tables used
        relations=[]            #temporarily stores the name of the schema and table used in the datasource
        #connection=[]           
        relations_tup=[]        #array of arrays - where each array has the name of schema and table

        for i in datasource.iter('relation'):  #for storing all the schema and table names in relations_tup
            if(i.get('table')!=None and i.get('table').find('.')!=-1):
                schema=i.get('table').split(".")[0].strip('[]')
                tab=i.get('table').split(".")[1].strip('[]')
                relations.append(tab)
                relations.append(schema)
                relations_tup.append(relations)
                relations=[]

        dsn_tns = cx_Oracle.makedsn(hostnameEntry.get(), portEntry.get(), service_name=serviceEntry.get())
        conn = cx_Oracle.connect(user=e1.get(), password=e2.get(), dsn=dsn_tns)         #Oracle DB connection
        cur=conn.cursor()
        
        '''
        for i in datasource.iter('connection'):
            if(i.get('schema')!=None):
                connection.append(i.get('schema'))
        #print(connection)
        '''
        
        for i in datasource.iter('column'):     #stores all column tags in columns array
            columns.append(i)

        for ele in columns:     #getting matching columns and secondary tables
            if(ele.get('name').strip('[]').upper().find(" ") != -1):
                match_col.append(ele.get('name').strip('[]').upper().split()[0])
                tableName_sec.append(ele.get('name').strip('[]').upper().split()[1].strip('()'))
                
        for t in relations_tup:
            #print(t)
            query= "SELECT * FROM dba_col_comments WHERE TABLE_NAME=:tab and OWNER=:ow"     #Querying the data dictionary
            cur.execute(query,[t[0],t[1]])
            
            colName=[]      #cursor column names
            tableName=[]    #cursor table names
            comments=[]     #cursor comments
            
            for row in cur:                 #filling in the above cursor arrays
                tableName.append(row[1])
                colName.append(row[2])
                comments.append(row[3])
                
            for i in range(0,len(tableName)):           #getting the primary table name
                if(tableName[i] not in tableName_sec):
                    primary.append(tableName[i])
                    break
                
            for ele in columns:
                #print(ele.get('name'))
                for i in range(0,len(colName)):
                    if(ele.get('name').strip('[]').upper().find(" ") != -1):        #for adding descriptions of matching columns
                        if(colName[i] == ele.get('name').strip('[]').upper().split()[0] and tableName[i]==ele.get('name').strip('[]').upper().split()[1].strip('()')):
                            #print(colName[i],tableName[i],"c1")
                            if(ele.find('desc')!=None):     #updates the description if 'desc' tag already present
                                desc=ele.find('desc')
                                ft=desc.find('formatted-text')
                                run=ft.find('run')
                                run.text=comments[i]
                            else:                           #creates desc tag and adds the description
                                desc=ET.Element('desc')
                                ele.append(desc)
                                ft=ET.Element('formatted-text')
                                desc.append(ft)
                                run=ET.Element('run')
                                ft.append(run)
                                run.text=comments[i]
                            break
                    elif(colName[i] == ele.get('name').strip('[]').upper() and tableName[i] in primary):       #for adding descriptions to columns of primary tables 
                        #print(colName[i],tableName[i],"c2")
                        if(ele.find('desc')!=None):
                            desc=ele.find('desc')
                            ft=desc.find('formatted-text')
                            run=ft.find('run')
                            run.text=comments[i]
                        else:
                            desc=ET.Element('desc')
                            ele.append(desc)
                            ft=ET.Element('formatted-text')
                            desc.append(ft)
                            run=ET.Element('run')
                            ft.append(run)
                            run.text=comments[i]
                        break    
                    else:
                        if(colName[i] == ele.get('name').strip('[]') and colName[i] not in match_col):      #for adding description to rest of the columns
                            #print(colName[i],tableName[i],"c3")
                            if(ele.find('desc')!=None):
                                desc=ele.find('desc')
                                ft=desc.find('formatted-text')
                                run=ft.find('run')
                                run.text=comments[i]
                            else:
                                desc=ET.Element('desc')
                                ele.append(desc)
                                ft=ET.Element('formatted-text')
                                desc.append(ft)
                                run=ET.Element('run')
                                ft.append(run)
                                run.text=comments[i]
                            break

        tree.write(e.get())
       
    elif(custom==1):
        tree = ET.parse(e.get())
        root = tree.getroot()
        datasources = root.find('datasources')
        datasource = datasources.find('datasource')

        query=""            #stores custom SQL query
        columnNames=[]      #stores column names
        datatypes=[]        #stores datatypes of the related columns
        relations=[]        #stores all the tables used in the query
        primary=[]          #stores the table name which was first connected
        secondary=[]        #stores the table names which were connected after primary table
        match_col=[]        #stores matching column names

        for i in datasource.iter('relation'):   #getting the query from 'relation' tag
            query=i.text
            break
        '''print("query: ",query)
        print('\n')'''

        for i in datasource.iter('local-name'):     #filling in column names
            columnNames.append(i.text)
        '''print("columnNames: ",columnNames)
        print('\n')'''

        for i in datasource.iter('local-type'):     #filling in datatypes
            datatypes.append(i.text)
        '''print("datatypes: ",datatypes)
        print('\n')'''

        tableNames=sql_metadata.get_query_tables(query)
        for i in tableNames:
            relations.append(i.split("."))
        #print("relations: ",relations)
        primary.append(relations[0][1].upper())
        #print("primary: ",primary)
        #print('\n')
        for i in range(1,len(relations)):
            secondary.append(relations[i][1].upper())       #filling in the relations, primary and secondary array
        #print("secondary: ",secondary)
        #print('\n')

        dsn_tns = cx_Oracle.makedsn(hostnameEntry.get(), portEntry.get(), service_name=serviceEntry.get())
        conn = cx_Oracle.connect(user=e1.get(), password=e2.get(), dsn=dsn_tns)         #Oracle DB connection
        cur=conn.cursor()           #Oracle DB connection

        for ele in columnNames:             #getting matching columns
            if(ele.strip('[]').upper().find(" ") != -1):
                match_col.append(ele.strip('[]').upper().split()[0])
        #print("match_col: ",match_col)

        for t in relations:
            #print("\n")
            #print("t[0] t[1]: ",t[0],t[1])
            query= "SELECT * FROM dba_col_comments WHERE TABLE_NAME=:tab and OWNER=:ow"     #Querying the data dictionary
            cur.execute(query,[t[1].upper(),t[0].upper()])
            
            cur_colName=[]      #cursor column names
            cur_tableName=[]    #cursor table names
            comments=[]         #cursor comments
            
            for row in cur:     #filling in the above cursor arrays
                cur_tableName.append(row[1])
                cur_colName.append(row[2])
                comments.append(row[3])
            '''print('cursor colname: ',cur_colName)
            print('\n')
            print('cursor tableNames: ',cur_tableName)
            print('\n')
            print('comments: ',comments)
            print('\n')'''
            
            flag=0  #indicates the existence of the respective column's tag(0 means not present and 1 indicates that the tag is present)
            
            for ele in range(0,len(columnNames)):
                for i in range(0,len(cur_colName)):
                    flag=0
                    if(columnNames[ele].strip('[]').find(" ")!=-1):
                        if(columnNames[ele].strip('[]').split()[0]==cur_colName[i] and cur_tableName[i].find(columnNames[ele].strip('[]').split()[1].strip('()'))!=-1):
                            #print(columnNames[ele],cur_colName[i],"c1")
                            for ds in datasource.iter('column'):
                                if(ds.get('name')==columnNames[ele] and ds.get('semantic-role')==None):
                                    #print(columnNames[ele],cur_colName[i],"c11")
                                    desc=ds.find('desc')
                                    ft=desc.find('formatted-text')
                                    run=ft.find('run')
                                    run.text=comments[i]
                                    flag=1
                                    break
                            if(flag==0):
                                #print(columnNames[ele],cur_colName[i],"c12")
                                column=ET.Element('column')
                                column.set('datatype',datatypes[i])
                                column.set('name',columnNames[ele])
                                if(datatypes[ele]=='real' or datatypes[ele]=='integer'):
                                    column.set('role','measure')
                                    column.set('type','quantitative')
                                elif(datatypes[ele]=='datetime' or datatypes[ele]=='date'):
                                    column.set('role','dimension')
                                    column.set('type','ordinal')
                                else:
                                    column.set('role','dimension')
                                    column.set('type','nominal')
                                datasource.insert(1,column)
                                desc=ET.Element('desc')
                                column.append(desc)
                                ft=ET.Element('formatted-text')
                                desc.append(ft)
                                run=ET.Element('run')
                                ft.append(run)
                                run.text=comments[i]
                                break
                    elif((cur_colName[i]==columnNames[ele].strip('[]') and cur_tableName[i] in primary)):
                        #print(columnNames[ele],cur_colName[i],"c2")
                        for ds in datasource.iter('column'):
                            if(ds.get('name')==columnNames[ele] and ds.get('semantic-role')==None):
                                #print(columnNames[ele],cur_colName[i],"c21")
                                desc=ds.find('desc')
                                ft=desc.find('formatted-text')
                                run=ft.find('run')
                                run.text=comments[i]
                                flag=1
                                break
                        if(flag==0):  
                            #print(columnNames[ele],cur_colName[i],"c22")
                            column=ET.Element('column')
                            column.set('datatype',datatypes[i])
                            column.set('name',columnNames[ele])
                            if(datatypes[ele]=='real' or datatypes[ele]=='integer'):
                                column.set('role','measure')
                                column.set('type','quantitative')
                            elif(datatypes[ele]=='datetime' or datatypes[ele]=='date'):
                                column.set('role','dimension')
                                column.set('type','ordinal')
                            else:
                                column.set('role','dimension')
                                column.set('type','nominal')
                            datasource.insert(1,column)
                            desc=ET.Element('desc')
                            column.append(desc)
                            ft=ET.Element('formatted-text')
                            desc.append(ft)
                            run=ET.Element('run')
                            ft.append(run)
                            run.text=comments[i]
                            break
                    elif(cur_colName[i]==columnNames[ele].strip('[]') and cur_colName[i] not in match_col):
                        #print(columnNames[ele],cur_colName[i],"c3")
                        for ds in datasource.iter('column'):
                            if(ds.get('name')==columnNames[ele] and ds.get('semantic-role')==None):
                                #print(columnNames[ele],cur_colName[i],"c31")
                                desc=ds.find('desc')
                                ft=desc.find('formatted-text')
                                run=ft.find('run')
                                run.text=comments[i]
                                flag=1
                                break
                        if(flag==0):
                            #print(columnNames[ele],cur_colName[i],"c32")
                            column=ET.Element('column')
                            column.set('datatype',datatypes[i])
                            column.set('name',columnNames[ele])
                            if(datatypes[ele]=='real' or datatypes[ele]=='integer'):
                                column.set('role','measure')
                                column.set('type','quantitative')
                            elif(datatypes[ele]=='datetime' or datatypes[ele]=='date'):
                                column.set('role','dimension')
                                column.set('type','ordinal')
                            else:
                                column.set('role','dimension')
                                column.set('type','nominal')
                            datasource.insert(1,column)
                            desc=ET.Element('desc')
                            column.append(desc)
                            ft=ET.Element('formatted-text')
                            desc.append(ft)
                            run=ET.Element('run')
                            ft.append(run)
                            run.text=comments[i]
                            break
                                
        tree.write(e.get())    #writing to the XML workbook
    tkinter.messagebox.showinfo("","Done! Reopen your workbook")    #Completion messagebox

#Tkinter GUI   

root = themes.ThemedTk()     
root.get_themes()
root.set_theme("breeze")
root.geometry("800x700")
root.title("Tableau - Import column descriptions v1.0")

img = Image.open('mchpLogo.jpg') 
img = img.resize((150, 150), Image.ANTIALIAS)
img = ImageTk.PhotoImage(img) 
panel = Label(root, image = img) 
panel.pack(padx=5, pady=5) 


L4=ttk.Label(root, text="Hostname : ")
L4.pack(padx=5, pady=5)
hostnameEntry=ttk.Entry(root,width=60)
hostnameEntry.pack(padx=5, pady=5)
L5=ttk.Label(root, text="Port : ")
L5.pack(padx=5, pady=5)
portEntry=ttk.Entry(root,width=60)
portEntry.pack(padx=5, pady=5)
L6=ttk.Label(root, text="Service name : ")
L6.pack(padx=5, pady=5)
serviceEntry=ttk.Entry(root,width=60)
serviceEntry.pack(padx=5, pady=5)

L1=ttk.Label(root, text="Username : ")
L1.pack(padx=5, pady=5)
e1=ttk.Entry(root,width=60)
e1.pack(padx=5, pady=5)
L2=ttk.Label(root, text="Password : ")
L2.pack()
e2=ttk.Entry(root,width=60,show='*')
e2.pack(padx=5, pady=5)
L=ttk.Label(root, text="Enter your workbook name : ")
L.pack()
e=ttk.Entry(root,width=60)
e.pack(padx=5, pady=5)
browseButton=ttk.Button(root,text="Browse file",command=browse)
browseButton.pack(padx=5, pady=5)
myButton=ttk.Button(root,text="Submit",command = lambda:[waitingMsg(), myClick(), waitingMsg()])
myButton.pack(padx=5, pady=5)
L3=ttk.Label(root, text="", )
L3.pack()
root.mainloop()