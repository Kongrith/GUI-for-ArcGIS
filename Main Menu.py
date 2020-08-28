import sys,os
import psycopg2
import traceback

from PyQt4 import QtGui
from PyQt4 import QtCore

ar1 =  "[% "STATIONID" %]"
ar1 = ar1.rstrip()

try:
    # this:
    conn = psycopg2.connect(database="DisCon", user="postgres", password="1234", host="localhost", port= 5432)
    # print "connected"

except psycopg2.Error as e:
    print "I am unable to connect to the database"

cur = conn.cursor()
cur.execute("SELECT subid, eng, thai, swid  from idtable")
IDTable = cur.fetchall()

ChkSw = False
for row in IDTable:
   if (ar1== row[0].rstrip()):                                                                         #  ถ้าข้อมูลในฐานข้อมูลตรงกับรหัสย่อสถานีไฟฟ้า
         MenuSubTitle =  row[1].rstrip() + ' '+ "Substation Informations"       #  สร้างชื่อเต็มของสถานีไฟฟ้า
         IDnum = str(row[3])                                                                            #  เก็บค่า switching number เพื่อติดต่อกับเวป อคฟ.
         print IDnum
         if IDnum != 'None' :
                 WebPath =  "http://control.egat.co.th/smots/edit/downloadfile.aspx?type=sw&sub=" + IDnum
                 print MenuSubTitle
                 ChkSw = True 
         else:
                 ChkSw = False         
cur.close()  
conn.close()

class MessageBox(QtGui.QWidget):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)

        self.resize(480, 300)
        self.setWindowTitle(MenuSubTitle)

        button1 = QtGui.QPushButton("Transformer and Line Loading", self)
        button1.move(20,20)
        button1.resize(180,30)
        button2 = QtGui.QPushButton("Long-Term Substation Load Forecast", self)
        button2.move(20,70)
        button2.resize(180,30)
        button3 = QtGui.QPushButton("Switching Diagram", self)
        button3.move(20,120)
        button3.resize(180,30)
        button4 = QtGui.QPushButton(" Single Line Diagram ", self)
        button4.move(20,170)
        button4.resize(180,30)
        button5 = QtGui.QPushButton("General Arrangement Plan", self)
        button5.move(20,220)       
        button5.resize(180,30)

        button1.clicked.connect(self.ActualLoad)
        button2.clicked.connect(self.ForecastLoad)
        button3.clicked.connect(self.SW)
        button4.clicked.connect(self.SLD)
        button5.clicked.connect(self.GeneralArrangement)

    def ActualLoad(self):
        QtGui.QMessageBox.information(None, "Feature id", "feature id is[% "STATIONID" %]")

    def ForecastLoad(self):
        QtGui.QMessageBox.information(None, "Feature id", "feature id is[% "STATIONID" %]")

    def SW(self):
        import webbrowser
        if ChkSw == True:
            chrome_path = 'C:\Program Files\Google\Chrome\Application\chrome.exe %s'
            webbrowser.get('windows-default').open_new_tab(WebPath)
        else:
            print "can't retrieve data from http://control.egat.co.th !"

    def SLD(self):
        directory = "D:/DataGIS/Avenger/" + ar1
        ChkSLD = False
        for root, dirs, files in os.walk(directory):
            filename = os.path.join(files)
        for i in range(0,len(filename)):
            if filename[i] == ar1 + ' SL 115.pdf':                          # สำหรับกรณี ex. LE SL 115.pdf   (อันนี้เป็นมาตราน)
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + ' SL 115.pdf'        
                 os.startfile(SLDpath)           
            if filename[i] == 'SL 115.pdf':                                 # สำหรับกรณี ex. SL 115.pdf 
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 115.pdf'        
                 os.startfile(SLDpath)
            if filename[i] == 'SL 115.tif':                                 # สำหรับกรณี ex. SL 115.tif 
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 115.tif'        
                 os.startfile(SLDpath)
            if filename[i] ==  'SL 115 ' + ar1 + '.pdf':                    # สำหรับกรณี ex. SL 115 LE.pdf  
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + 'SL 115 ' + ar1 + '.pdf'        
                 os.startfile(SLDpath)            
            if filename[i] ==  'SL 115 ' + ar1 + '.tif':                    # สำหรับกรณี ex. SL 115 LE.tif  
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + 'SL 115 ' + ar1 + '.tif'        
                 os.startfile(SLDpath)
            
            if filename[i] == ar1 + ' SL 230.pdf':                          # สำหรับกรณี ex. LE SL 230.pdf   (อันนี้เป็นมาตราน)
                 ChkSLD = True
                 print('my computer showed 230kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + ' SL 230.pdf'        
                 os.startfile(SLDpath)            
            if filename[i] == 'SL 230.pdf':                                 # สำหรับกรณี ex. SL 230.pdf 
                 ChkSLD = True
                 print('my computer showed 230kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 230.pdf'        
                 os.startfile(SLDpath)
            if filename[i] ==  'SL 230 ' + ar1 + '.pdf':                    # สำหรับกรณี ex. SL 115 LE.pdf  
                 ChkSLD = True
                 print('my computer showed 230kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + 'SL 230 ' + ar1 + '.pdf'        
                 os.startfile(SLDpath)            
            if filename[i] == 'SL 230.tif':                                 # สำหรับกรณี ex. SL 230.tif 
                 ChkSLD = True
                 print('my computer showed 230kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 230.tif'        
                 os.startfile(SLDpath)            

            if filename[i] == ar1 + ' SL 500.pdf':                          # สำหรับกรณี ex. LE SL 500.pdf   (อันนี้เป็นมาตราน)
                 ChkSLD = True
                 print('my computer showed 500kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + ' SL 500.pdf'        
                 os.startfile(SLDpath)            
            if filename[i] == 'SL 500.pdf':                                 # สำหรับกรณี ex. SL 500.pdf 
                 ChkSLD = True
                 print('my computer showed 500kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 500.pdf'        
                 os.startfile(SLDpath)
            if filename[i] == 'SL 500.tif':                                 # สำหรับกรณี ex. SL 500.tif 
                 ChkSLD = True
                 print('my computer showed 115kV Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/SL 500.tif'        
                 os.startfile(SLDpath)

            if filename[i] == ar1 + '-S-1.pdf':                             # สำหรับกรณี ex. LE-S-1.pdf   (common)
                 ChkSLD = True
                 print('my computer showed Single Line Diagram !')
                 SLDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + '-S-1.pdf'        
                 os.startfile(SLDpath)
            
        if ChkSLD == False:
            print('my computer not found any Single Line Diagram !')

    def GeneralArrangement(self):		
        directory = "D:/DataGIS/Avenger/" + ar1
        ChkLO = False
        for root, dirs, files in os.walk(directory):
            filename = os.path.join(files)
        for i in range(0,len(filename)):
            if filename[i] == ar1 + ' LO.pdf':                             # สำหรับกรณี ex. LE LO.pdf
                 ChkLO = True
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + ' LO.pdf'        
                 os.startfile(LOpath)        
            if filename[i] == '1.jpg':                                      # สำหรับกรณี ex. 1.jpg
                 ChkLO = True 
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/'+ '1.jpg'
                 os.startfile(LOpath)
            if filename[i] == 'LO.pdf':                                      # สำหรับกรณี ex. LO.pdf
                 ChkLO = True 
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/'+ 'LO.pdf'
                 os.startfile(LOpath)            
            if filename[i] == 'LO'+ " " + ar1 + '.pdf':                      # สำหรับกรณี ex. LO LE.pdf
                 ChkLO = True
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/LO '+ ar1 + '.pdf'        
                 os.startfile(LOpath)
            if filename[i] == 'LO'+ "-" + ar1 + '.pdf':                      # สำหรับกรณี ex. LO-LE.pdf
                 ChkLO = True
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/LO-'+ ar1 + '.pdf'        
                 os.startfile(LOpath)
            if filename[i] == 'Location'+ "_" + ar1 + '.pdf':                # สำหรับกรณี ex. Location_LE.pdf
                 ChkLO = True
                 print('my computer showed location !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/Location_'+ ar1 + '.pdf'        
                 os.startfile(LOpath)
            if filename[i] == ar1 + '-LO.pdf':                               # สำหรับกรณี ex. LE-LO.pdf
                 ChkLO = True
                 print('my computer showed location !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/'+ ar1 + '-LO.pdf'        
                 os.startfile(LOpath)            
            if filename[i] == 'LO.tif':                                      # สำหรับกรณี ex. LO.tif
                 ChkLO = True
                 print('my computer showed Substation Layout !')
                 LOpath = "D:/DataGIS/Avenger/" + ar1 + '/'+ 'LO.tif'
                 os.startfile(LOpath)
            
            if filename[i] == ar1 + '-S-2.pdf':                              # สำหรับกรณี ex. LE-S-2.pdf   (common)
                 ChkSLD = True
                 print('my computer showed Substation Layout !')
                 LDpath = "D:/DataGIS/Avenger/" + ar1 + '/' + ar1 + '-S-2.pdf'        
                 os.startfile(SLDpath)
            
        if ChkLO == False:
            print('my computer not found any Substation Layout !')
				
    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Message',"Are you sure to quit?", QtGui.QMessageBox.Yes,QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()            
        else:
            event.ignore()

qb = MessageBox()
qb.show()
