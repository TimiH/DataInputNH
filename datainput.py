#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
from PyQt4 import QtCore, QtGui, uic
import pandas as pd
from openpyxl import load_workbook
import os.path
from collections import OrderedDict

fnameMammo = os.path.join(os.path.curdir,'data',"DataMammo.xlsx")
fnameSono = os.path.join(os.path.curdir,'data',"DataSono.xlsx")
fnameMrt = os.path.join(os.path.curdir,'data',"DataMrt.xlsx")
qtCreatorFile = "ninaDatainput.ui" # Enter file here.


Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class MyApp(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        #connect Buttons
        self.saveMam.clicked.connect(self.saveMamData)
        self.saveMrt.clicked.connect(self.saveMrtData)
        #self.saveSono.clicked.connect(self.saveSono)

    def saveSonoData(self):
        if self.sonoCheck() == 0:
            self.createSono()


    def saveMamData(self):
        #check if file exists
        if self.mamCheck() == 0:
            self.createMam()

        #get the Data
        #name etc
        vname = str(self.vname.text())
        nname = str(self.nname.text())
        studyid =  str(self.studyid.text())
        histo =  str(self.histo.text())
        #date intern
        dateMam = str(self.dateMam.date().toString("dd.MM.yy"))
        internMam = int(self.internMam.isChecked())
        #birads and Parent
        biradLeft = self.biradLeftMam.currentIndex() #if none is selected bottom index =0
        biradRight = self.biradRightMam.currentIndex()
        parenLeft = self.parenLeftMam.currentIndex()
        parenRight = self.parenRightMam.currentIndex()
        #Befund
        herdMam = int(self.herdMam.isChecked())
        zweitMam = int(self.zweitMam.isChecked())
        laterMam = self.laterMam.currentIndex()
        quadMam = self.quadMam.currentIndex()
        depthMam = self.depthMam.currentIndex()
        formMam = self.formMam.currentIndex()
        randMam = self.randMam.currentIndex()
        dichteMam = self.dichteMam.currentIndex()
        sizeMam =  str(self.nname.text())
        #Verkalkung
        kalkMam = int(self.kalkMam.isChecked())
        typischMam = int(self.typischMam.isChecked())
        typischSusMam = self.typischSusMam.currentIndex()
        verteilungMam = self.verteilungMam.currentIndex()
        #Weiteres
        archMam = int(self.archMam.isChecked())
        intraMam = int(self.intraMam.isChecked())
        asymmMam = self.asymmMam.currentIndex()
        #Begleitmerkmale
        hautretraktionMam = int(self.hautretraktionMam.isChecked())
        mamillenMam = int(self.mamillenretraktionMam.isChecked())
        cutisMam = int(self.cutisMam.isChecked())
        trabelMam = int(self.trabelMam.isChecked())
        axillMam = int(self.axillMam.isChecked())
        archBMam = int(self.archBMam.isChecked())
        kalkBMam = int(self.kalkBMam.isChecked())

        #TODO
        #Logic to stop unwanted datainput
            #if herdläsion is unchecked ignore all inputs inside Befund
            #if Verkalkung is unchecked do what??
            #if any of the first 4 is empty raise alarm


        #append to list
        values = []
        values.append(studyid)
        values.append(vname)
        values.append(nname)
        values.append(histo)
        values.append(dateMam)
        values.append(internMam)
        #biradsParen
        values.append(biradLeft)
        values.append(biradRight)
        values.append(parenLeft)
        values.append(parenRight)
        #Befund
        values.append(herdMam)
        values.append(zweitMam)
        values.append(laterMam)
        values.append(quadMam)
        values.append(depthMam)
        values.append(formMam)
        values.append(randMam)
        values.append(dichteMam)
        values.append(sizeMam)
        #Verkalkung
        values.append(kalkMam)
        values.append(typischMam)
        values.append(typischSusMam)
        values.append(verteilungMam)
        #Weiteres
        values.append(archMam)
        values.append(asymmMam)
        values.append(intraMam)
        #Begleitmerkmale
        values.append(hautretraktionMam)
        values.append(mamillenMam)
        values.append(cutisMam)
        values.append(trabelMam)
        values.append(axillMam)
        values.append(archBMam)
        values.append(kalkMam)

        #append to datafile in /data/DataMammo #TODO ORDERCheck
        df = pd.read_excel(fnameMammo,index=False,encoding='utf-8')
        headers = df.columns.tolist()
        dictTemp = OrderedDict(zip(headers,values))
        newRow = pd.DataFrame([dictTemp])
        df = df.append(newRow)
        df.to_excel(fnameMammo, index=False, header = 1,encoding='utf-8')

    #TODO Refactor check,create

    def mamCheck(self):
        #the //Data folder needs to be existent
        fname = os.path.join(os.path.curdir,'data', fnameMammo)
        return os.path.isfile(fname)

    def sonoCheck(self):
        #the //Data folder needs to be existent
        fname = os.path.join(os.path.curdir,'data', fnameSono)
        return os.path.isfile(fname)

    def mrtCheck(self):
        #the //Data folder needs to be existent
        fname = os.path.join(os.path.curdir,'data', fnameMrt)
        return os.path.isfile(fname)

    def createMam(self):
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'mammo.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMama.xlxs"
        df.to_excel(fnameMammo, index=False)

    def createSono(self):
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'sono.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMama.xlxs"
        df.to_excel(fnameSono, index=False)

    def createMrt(self):
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'mrt.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMama.xlxs"
        df.to_excel(fnameMrt, index=False)


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
