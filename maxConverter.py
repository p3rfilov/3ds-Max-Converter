
import sys, re, win32api, glob, os, subprocess, signal, time, threading
from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QTableWidgetItem, QHeaderView, QMessageBox
from PyQt5.uic import loadUi

'''
TABLE STRUCTURE
ROW - INFO

0 - File Name
1 - File Version
2 - Status
3 - Path
4 - Data
'''

def getFileList(folder): #get all files in folder
    if folder:
        fList = []
        os.chdir(folder)
        for file in glob.glob("*.max"):
            fList.append(file)
        return fList


class mainWindow(QDialog):
    
    def __init__(self):
        QDialog.__init__(self)
        self.ui = loadUi('maxConverterGUI.ui')
        self.ui.show()

        self.maxPaths = ['Program Files\\Autodesk\\3ds*','Program Files (x86)\\Autodesk\\3ds*']
        self.canConvertTo = []
        self.int_ConvertTo = []
        self.Installs = []
        self.installDict = {}
        self.process = ''
        self.stop = False
        self.watchList = []
        self.barCount = 0
        
        # drag and drop attempt
        #table = self.ui.tbl_fileList
        #table.setAcceptDrops(True)
        #table.dropped.connect(self.filesDropped)
        # drag and drop attempt
        
        action = self.ui.menuHelp.addAction('&About')
        action.triggered.connect(self.About)
        
        self.installDict = self.maxInstalls(self.maxPaths)
        self.Installs = sorted(int(key) for key in self.installDict)
        strInstalls = ', '.join(str(i) for i in self.Installs)
        for i in self.Installs: self.buildConvertList(i, self.canConvertTo)
        self.canConvertTo = sorted(list(set(self.canConvertTo)))
        self.int_ConvertTo = [int(i) for i in self.canConvertTo]
        
        self.ui.btn_Add.clicked.connect(self.addFiles)
        self.ui.btn_Remove.clicked.connect(self.removeFiles)
        self.ui.cbx_convertTo.currentIndexChanged.connect(self.verSelected)
        self.ui.btn_Convert.clicked.connect(self.makeBatch)
        self.ui.btn_Stop.clicked.connect(self.stopScript)

        self.ui.btn_Convert.setEnabled(False)
        self.ui.tbl_fileList.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.ui.tbl_fileList.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.ui.tbl_fileList.setColumnHidden(3,True)
        self.ui.tbl_fileList.setColumnHidden(4,True)
        self.ui.lbl_maxVers.setText(strInstalls)
        self.ui.cbx_convertTo.clear()
        for i in reversed(self.canConvertTo): self.ui.cbx_convertTo.addItem(i)

#     @QtCore.pyqtSlot(list)
#     def filesDropped(self, links):
#         print('got signal')
#         print(links)

    def About(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(
            '''
            -------------------------------------------------
            
            3ds Max File Converter v1.2
            by Andrej Perfilov
            
            Constructed using Python 3.5 and Qt 5.6
            
            -------------------------------------------------
            '''
            )
        msg.setDetailedText(
            '''
        Please note that you still need all 
        the required installations of 3ds Max 
        to convert the files.
            
        Supported 3ds Max versions:
            
        2011 : 2010, 2011
        2012 : 2010, 2011, 2012   
        2013 : 2010, 2011, 2012, 2013
        2014 : 2011, 2012, 2013, 2014
        2015 : 2012, 2013, 2014, 2015
        2016 : 2012, 2013, 2014, 2015, 2016
        2017 : 2014, 2015, 2016, 2017
            '''
            )
        msg.setWindowTitle('About')
        msg.exec_()
        msg.setWindowModality(QtCore.Qt.ApplicationModal)


    def maxInstalls(self, pathList): # search for 3ds Max installations
        if pathList:
            exe = '\\3dsmax.exe'
            exePaths = []
            versions = []
            drives = win32api.GetLogicalDriveStrings()
            drives = drives.split('\000')[:-1]
            for d in drives:
                for f in pathList:
                    maxEXE = glob.glob(d + f + exe)
                    if maxEXE:
                        exePaths += maxEXE
            for i in exePaths:
                ver = os.path.dirname(i)[-5:] #last 5 characters of the 3dsmax.exe location directory
                ver = re.sub('[^0123456789]', '', ver) #getting version number
                versions.append(ver)
            return dict(zip(versions, exePaths))


    def buildConvertList(self, ver, theList): # find out which versions we can covert to
        if ver:
            case = int(ver)
            if case < 2011:
                pass
            elif case == 2011:
                theList.extend(('2010','2011'))
            elif case == 2012:
                theList.extend(('2010','2011','2012'))
            elif case == 2013:
                theList.extend(('2010','2011','2012','2013'))
            elif case == 2014:
                theList.extend(('2011','2012','2013','2014'))
            elif case == 2015:
                theList.extend(('2012','2013','2014','2015'))
            elif case == 2016:
                theList.extend(('2012','2013','2014','2015','2016'))
            elif case == 2017:
                theList.extend(('2014','2015','2016','2017'))    
            elif case > 2017:
                theList.extend((str(case),))
            return theList


    def fileVer(self, maxFile): # check .max file version
        savedVer = 'SavedAsVersion:'
        maxVer = '3dsmaxVersion:'
        if maxFile:
            def getVersion(st, line):
                i = line.index(st)+len(st) #getting string position
                ver = line[i:i+5]
                ver = re.sub('[^0123456789.]', '', ver) #getting version number
                ver = ver.rstrip('0')
                ver = ver.rstrip('.')
                ver = int(ver)
                if ver > 9:
                    ver += 1998 #setting 3ds Max version
                return str(ver)
            def reverse_readline(filename): # big thanks to srohde and Andomar at http://stackoverflow.com for this piece of code!
                """a generator that returns the lines of a file in reverse order"""
                buf_size=8192
                with open(filename, 'rt', encoding='latin-1') as fh:
                    segment = None
                    offset = 0
                    fh.seek(0, os.SEEK_END)
                    file_size = remaining_size = fh.tell()
                    while remaining_size > 0:
                        offset = min(file_size, offset + buf_size)
                        fh.seek(file_size - offset)
                        buffer = fh.read(min(remaining_size, buf_size))
                        remaining_size -= buf_size
                        lines = buffer.split('\n')
                        # the first line of the buffer is probably not a complete line so
                        # we'll save it and append it to the last line of the next buffer
                        # we read
                        if segment is not None:
                            # if the previous chunk starts right from the beginning of line
                            # do not concact the segment to the last line of new chunk
                            # instead, yield the segment first 
                            if buffer[-1] is not '\n':
                                lines[-1] += segment
                            else:
                                yield segment
                        segment = lines[0]
                        for index in range(len(lines) - 1, 0, -1):
                            if len(lines[index]):
                                yield lines[index]
                    # Don't yield None if the file was empty
                    if segment is not None:
                        yield segment
            for line in reverse_readline(maxFile):
                cleanLine = re.sub('[^AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789.:]', '', line) #filtering out unneeded characters
                if cleanLine:
                    if savedVer in cleanLine:
                        return getVersion(savedVer, cleanLine)
                    elif maxVer in cleanLine:
                        return getVersion(maxVer, cleanLine)


    def addFiles(self):
        table = self.ui.tbl_fileList
        names = QFileDialog.getOpenFileNames(self, 'Open files', '*.max')
        if names[0]:
            self.ui.btn_Remove.setEnabled(True)
            self.ui.btn_Convert.setEnabled(True)
            paths = [table.item(j,3).text() for j in range(table.rowCount())] # gather file paths from table
            for i in names[0]:
                if i not in paths:
                    self.addToTable(i)
            for i in range(table.rowCount()):
                if table.item(i,1).text() == 'Processing...':
                    ver = self.fileVer(table.item(i,3).text())
                    if ver == None:
                        table.item(i,1).setText('Unknown')
                    else:
                        table.item(i,1).setText(ver)
                        table.item(i,2).setText('Ready')
            self.verSelected()
                    #self.fList.insert(0,i)
                    #try: self.changeStatus()
                    #except: pass
                    #t = threading.Thread(target=self.getFileVer, args=(i,))
                    #t.start()

    def addToTable(self, file):
        table = self.ui.tbl_fileList
        if file:
            table.insertRow(0)
            i1 = QTableWidgetItem(os.path.basename(file))   # name
            i2 = QTableWidgetItem('Processing...')          # version
            i3 = QTableWidgetItem('')                       # status
            i4 = QTableWidgetItem(file)                     # path
            i5 = QTableWidgetItem('')                       # data
            i1.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
            i2.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
            i3.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
            i4.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
            i5.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
            i2.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
            i3.setTextAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter|QtCore.Qt.AlignCenter)
            table.setItem(0, 0, i1)
            table.setItem(0, 1, i2)
            table.setItem(0, 2, i3)
            table.setItem(0, 3, i4)
            table.setItem(0, 4, i5)

    def changeStatus(self):
        table = self.ui.tbl_fileList
        for row in range(table.rowCount()):
            if table.item(row, 2).text() == 'Ready':
                table.item(row, 1).setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                table.item(row, 2).setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                table.item(row, 2).setForeground(QtGui.QColor(0,0,0))
            else:
                table.item(row, 1).setFlags(QtCore.Qt.ItemIsSelectable)
                table.item(row, 2).setFlags(QtCore.Qt.ItemIsSelectable)
                table.item(row, 2).setForeground(QtGui.QColor(128,128,128))
        for row in range(table.rowCount()):
            if table.item(row, 2).text() == 'Converting...':
                table.item(row, 2).setForeground(QtGui.QColor(0,0,255))
            elif table.item(row, 2).text() == 'Converted!':
                table.item(row, 2).setForeground(QtGui.QColor(0,128,0))
            elif table.item(row, 2).text() == 'Stopped':
                table.item(row, 2).setForeground(QtGui.QColor(255,0,0))


    def removeFiles(self):
        table = self.ui.tbl_fileList
        index_list = []
        for model_index in table.selectionModel().selectedRows():
            index = QtCore.QPersistentModelIndex(model_index)
            index_list.append(index)
        for index in index_list:
            ind = index.row()
            table.removeRow(ind)
        if table.rowCount() < 1:
            self.ui.btn_Remove.setEnabled(False)
            self.ui.btn_Convert.setEnabled(False)
        self.verSelected()


    def getSteps(self,ver,target,installs,myList):
        A = int(ver) # file version as int
        B = target
        
        canConvertTo = []
        for i in installs:
            c = []
            self.buildConvertList(i,c)
            for j in c:
                canConvertTo.insert(0,j)
        canConvertTo = sorted(list(set(canConvertTo))) # list of all available conversions
        
        if str(target) in canConvertTo:
            if A > installs[-1]: # higher version than available
                myList.insert(0,'OOR') # Out of Range
            elif A == B:
                myList.insert(0,'SV') # same version
            elif A > B:
                revInst = list(reversed(installs))
                for i in revInst: # descending order
                    instList = []
                    self.buildConvertList(i,instList)
                    if str(B) in instList:
                        if str(A) in instList: # if both versions in list
                            myList.insert(0,str(i) + '-' + str(B))
                            break
                        else:
                            res = str(i) + '-' + str(B)
                            B = i
                            if res in myList: # if already in list stop recursion
                                myList.insert(0,'OOR') # out of range
                                break
                            else:
                                myList.insert(0,res)
                                self.getSteps(ver,B,installs,myList) # recursive function
            elif A < B:
                for i in installs: # ascending order
                    instList = []
                    instList = self.buildConvertList(i,instList)
                    if str(B) in instList:
                        myList.insert(0,str(i) + '-' + str(B))
                        break
        else:
            myList.insert(0,'OOR') # if not in canConvertTo: Out of range


    def verSelected(self):
        table = self.ui.tbl_fileList
        if self.ui.cbx_convertTo.currentText():
            for row in range(table.rowCount()):
                stepList = []
                ver = table.item(row, 1).text()
                target = int(self.ui.cbx_convertTo.currentText())
                if ver == 'Unknown':
                    table.item(row, 2).setText('Unavailable')
                elif int(ver) == target:
                    table.item(row, 2).setText('Same version')
                else:
                    self.getSteps(ver, target, self.Installs, stepList)
                    if 'OOR' in stepList:
                        table.item(row, 2).setText('Out of Range')
                    elif 'SV' in stepList:
                        table.item(row, 2).setText('Same version')
                    else:
                        table.item(row, 2).setText('Ready')
                        table.item(row, 4).setText(','.join(stepList))
        self.ui.p_bar.setValue(0)
        self.changeStatus()

    
    def makeBatch(self):
        self.ui.btn_Add.setEnabled(False)
        self.ui.btn_Remove.setEnabled(False)
        self.ui.btn_Convert.setEnabled(False)
        self.ui.cbx_convertTo.setEnabled(False)
        self.ui.btn_Stop.setEnabled(True)
        table = self.ui.tbl_fileList
        self.watchList = []
        self.barCount = 0
        wFail = False
        version = self.ui.cbx_convertTo.currentText()
        
        for row in range(table.rowCount()): # check if target file exists
            path = table.item(row, 3).text()
            target = os.path.splitext(path)[0] + '_max' + version + os.path.splitext(path)[1]
            if os.path.isfile(target):
                table.item(row, 2).setText('Converted!')
        self.changeStatus()
        
        for ins in reversed(self.Installs):
            f = ''
            msFile = r'C:\maxConvert_' + str(ins) + r'.ms'
            try:
                f = open(msFile, 'w') # create new or override existing
                f.close()
            except:
                if not wFail:
                    wFail = True
                    msg = QMessageBox()
                    msg.setIcon(QMessageBox.Critical)
                    msg.setText('\nFailed to create script file!\nMake sure you are running this program \"As Administrator\"      ')
                    msg.setWindowTitle('Read/Write Error')
                    msg.exec_()
                else: pass
            if f:
                files = []
                f = open(msFile, 'a')
                for row in range(table.rowCount()):
                    if table.item(row, 2).text() == 'Ready':
                        path = table.item(row, 3).text()
                        steps = table.item(row, 4).text().split(',') # get list of steps, ex: ['2017-2015','2015-2012']
                        if steps:
                            for i in steps:
                                A = i.split('-')[0]
                                B = i.split('-')[1]
                                if A == str(ins): # if target version == currently iterated install version
                                    if steps.index(i) == 0: A_file = path # if first iteration
                                    else: A_file = os.path.splitext(path)[0] + '_max' + A + os.path.splitext(path)[1]
                                    B_file = os.path.splitext(path)[0] + '_max' + B + os.path.splitext(path)[1]
                                    files.append(B_file)
                                    
                                    if A == B: # if target == destination, ex: 2017-2017
                                        f.write('loadMaxFile ' + r'"' + A_file + r'"' + ' quiet:true' + '\n')
                                        f.write('saveMaxFile ' + r'"' + B_file + r'"' + ' quiet:true' + '\n\n')
                                        #f.write('loadMaxFile ' + r'"' + A_file.replace('/',r'\\') + r'"' + ' quiet:true' + '\n')
                                        #f.write('saveMaxFile ' + r'"' + B_file.replace('/',r'\\') + r'"' + ' quiet:true' + '\n\n')
                                    else:
                                        f.write('loadMaxFile ' + r'"' + A_file + r'"' + ' quiet:true' + '\n')
                                        f.write('saveMaxFile ' + r'"' + B_file + r'"' + ' saveAsVersion:' + B + ' quiet:true' + '\n\n')
                                        #f.write('loadMaxFile ' + r'"' + A_file.replace('/',r'\\') + r'"' + ' quiet:true' + '\n')
                                        #f.write('saveMaxFile ' + r'"' + B_file.replace('/',r'\\') + r'"' + ' saveAsVersion:' + B + ' quiet:true' + '\n\n')
                                    self.barCount += 1
                                    QApplication.processEvents()
                f.write('quitMax #noPrompt')
                f.close()
                cmd = 'pushd ' + os.path.dirname(self.installDict[str(ins)]) + ' & ' + r'3dsmax.exe -q -U MAXScript ' + msFile
                self.watchList.append([cmd, files])
        self.runScript()
        self.stopScript()
        
                
    def runScript(self):
        table = self.ui.tbl_fileList
        version = self.ui.cbx_convertTo.currentText()
        self.stop = False
        count = 0
        
        for row in range(table.rowCount()):
            if table.item(row, 2).text() == 'Ready':
                table.item(row, 2).setText('Converting...')
        self.changeStatus()
        
        self.ui.p_bar.setValue(5)
        for i in self.watchList:
            if i[1]: # if there are files to watch
                if not self.stop:
                    self.process = subprocess.Popen(i[0], shell=True)
                    for j in i[1]:
                        while not os.path.isfile(j):
                            if self.stop: break
                            for _ in range(3):
                                time.sleep(1)
                                QApplication.processEvents()
                        if os.path.isfile(j):
                            for row in range(table.rowCount()):
                                path = table.item(row, 3).text()
                                target = os.path.splitext(path)[0] + '_max' + version + os.path.splitext(path)[1]
                                if j == target:
                                    table.item(row, 2).setText('Converted!')
                            count += 1
                            self.ui.p_bar.setValue(count/self.barCount*100)
                            self.changeStatus()
                            QApplication.processEvents()
    

    def stopScript(self):
        table = self.ui.tbl_fileList
        self.stop = True
        for row in range(table.rowCount()):
            if table.item(row, 2).text() == 'Converting...':
                table.item(row, 2).setText('Stopped')
        self.changeStatus()
        try:
            child_pid = self.process.pid
            os.kill(child_pid, signal.SIGTERM)
        except:
            print('Failed to stop child process')
        try:
            self.process.kill()
        except:
            print('Failed to stop main process')
            
        if self.ui.p_bar.value() != 100:
            self.ui.p_bar.setValue(0)
        self.ui.btn_Add.setEnabled(True)
        self.ui.btn_Remove.setEnabled(True)
        self.ui.btn_Convert.setEnabled(True)
        self.ui.cbx_convertTo.setEnabled(True)
        self.ui.btn_Stop.setEnabled(False)


if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = mainWindow()
    sys.exit(app.exec_())
