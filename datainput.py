#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys, os, subprocess
from PyQt4 import QtCore, QtGui, uic
import pandas as pd
from openpyxl import load_workbook
from collections import OrderedDict

fnameMammo = os.path.join(os.path.curdir,'data',"DataMammo.xlsx")
fnameSono = os.path.join(os.path.curdir,'data',"DataSono.xlsx")
fnameMrt = os.path.join(os.path.curdir,'data',"DataMrt.xlsx")
FILEPATH = os.path.abspath(__file__)
qtCreatorFile = "ninaDatainput.ui" # Enter file here.


Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

class MyApp(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)

        #connect Buttons
        self.saveMam.clicked.connect(self.saveMamData)
        self.saveSono.clicked.connect(self.saveSonoData)
        self.saveMrt.clicked.connect(self.saveMrtData)
        self.nextButton.clicked.connect(self.handleButton)
        #allow data input into Histologie Combobox
        self.histo.setEditable(True)

    def saveMrtData(self):
        #check if file exists
        if self.mrtCheck() == 0:
            self.createMrt()

        vname = unicode(self.vname.text()).encode("utf-8")
        nname = unicode(self.nname.text()).encode("utf-8")
        studyid = unicode(self.studyid.text()).encode("utf-8")
        histo = unicode(self.histo.currentText()).encode("utf-8")
        #date intern
        dateMrt = unicode(self.dateMrt.date().toString("dd.MM.yy"))
        internMrt = int(self.internMrt.isChecked())
        biradLeft = unicode(self.biradLeftMrt.currentText()).encode("utf-8") #if none is selected bottom index =0
        biradRight = unicode(self.biradRightMrt.currentText()).encode("utf-8")
        fggMrt = unicode(self.fggMrt.currentText()).encode("utf-8")
        strengthMrt = unicode(self.strengthMrt.currentText()).encode("utf-8")
        beidseitigMrt = unicode(self.beidseitigMrt.currentText()).encode("utf-8")
        #Befund
        herdMrt = int(self.herdMrt.isChecked())
        zweitMrt = int(self.zweitMrt.isChecked())
        laterMrt = unicode(self.laterMrt.currentText()).encode("utf-8")
        quadMrt = unicode(self.quadMrt.currentText()).encode("utf-8")
        depthMrt = unicode(self.depthMrt.currentText()).encode("utf-8")
        formMrt = unicode(self.formMrt.currentText()).encode("utf-8")
        randMrt = unicode(self.randMrt.currentText()).encode("utf-8")
        charakMrt = unicode(self.charakMrt.currentText()).encode("utf-8")
        sizeMrt = unicode(self.sizeMrt.text()).encode("utf-8")
        #Non Mass Enhancement
        verteilungMrt = unicode(self.verteilungMrt.currentText()).encode("utf-8")
        musterMrt = unicode(self.musterMrt.currentText()).encode("utf-8")
        sizeMassMrt = unicode(self.sizeMassMrt.text()).encode("utf-8")
        #Kinetic SI
        initialMrt = unicode(self.initialMrt.currentText()).encode("utf-8")
        lateMrt = unicode(self.lateMrt.currentText()).encode("utf-8")
        #Begleitmerkmale
        intraMrt = int(self.intraMrt.isChecked())
        mamillenMrt = int(self.mamillenMrt.isChecked())
        mamilleninfMrt = int(self.mamilleninfMrt.isChecked())
        hautMrt = int(self.hautMrt.isChecked())
        hautverMrt = int(self.hautverMrt.isChecked())
        hautinMrt = int(self.hautinMrt.isChecked())
        axilMrt = int(self.axilMrt.isChecked())
        pectoMrt = int(self.pectoMrt.isChecked())
        thoraxMrt = int(self.thoraxMrt.isChecked())
        archBMrt = int(self.archBMrt.isChecked())
        #Befunde ohne Enhancement
        zysteMrt = int(self.zysteMrt.isChecked())
        postopMrt = int(self.postopMrt.isChecked())
        posttheraMrt = int(self.posttheraMrt.isChecked())
        RaumMrt = int(self.RaumMrt.isChecked())
        archMrt = int(self.archMrt.isChecked())
        signalMrt = int(self.signalMrt.isChecked())

        #TODO
        #Logic to stop unwanted datainput

        #append to list
        values = []
        values.append(studyid)
        values.append(vname)
        values.append(nname)
        values.append(histo)
        values.append(dateMrt)
        values.append(internMrt)
        #biradsParen
        values.append(biradLeft)
        values.append(biradRight)
        values.append(fggMrt)
        values.append(strengthMrt)
        values.append(beidseitigMrt)
        #Befund
        values.append(herdMrt)
        values.append(zweitMrt)
        values.append(laterMrt)
        values.append(quadMrt)
        values.append(depthMrt)
        values.append(formMrt)
        values.append(randMrt)
        values.append(charakMrt)
        values.append(sizeMrt)
        #Non Mass Enhancement
        values.append(verteilungMrt)
        values.append(musterMrt)
        values.append(sizeMassMrt)
        #Kinetic SI
        values.append(initialMrt)
        values.append(lateMrt)
        #Begleitmerkmale
        values.append(intraMrt)
        values.append(mamillenMrt)
        values.append(mamilleninfMrt)
        values.append(hautMrt)
        values.append(hautverMrt)
        values.append(hautinMrt)
        values.append(axilMrt)
        values.append(pectoMrt)
        values.append(thoraxMrt)
        values.append(archBMrt)
        #Befunde ohne Enhancement
        values.append(zysteMrt)
        values.append(postopMrt)
        values.append(posttheraMrt)
        values.append(RaumMrt)
        values.append(archMrt)
        values.append(signalMrt)

        #append to datafile in /data/DataMammo #TODO ORDERCheck
        df = pd.read_excel(fnameMrt,index=False,encoding='utf-8')
        #print df
        headers = df.columns.tolist()
        dictTemp = OrderedDict(zip(headers,values))
        newRow = pd.DataFrame([dictTemp])
        #print newRow
        df = pd.concat([df, newRow])
        df.to_excel(fnameMrt, index=False, header = 1,encoding='utf-8')


    def saveSonoData(self):
        #check if file exists
        if self.sonoCheck() == 0:
            self.createSono()

        #get the Data
        #name etc
        vname = unicode(self.vname.text())
        nname = unicode(self.nname.text())
        studyid = unicode(self.studyid.text())
        histo = unicode(self.histo.currentText()).encode("utf-8")
        #date intern
        dateSono = unicode(self.dateSono.date().toString("dd.MM.yy"))
        internSono = int(self.internSono.isChecked())
        #birads and Parent
        biradLeft = unicode(self.biradLeftSono.currentText()).encode("utf-8") #if none is selected bottom index =0
        biradRight = unicode(self.biradRightSono.currentText()).encode("utf-8")
        parenLeft = unicode(self.parenLeftSono.currentText()).encode("utf-8")
        parenRight = unicode(self.parenRightSono.currentText()).encode("utf-8")
        #Befund
        herdSono = int(self.herdSono.isChecked())
        zweitSono = int(self.zweitSono.isChecked())
        laterSono = unicode(self.laterSono.currentText()).encode("utf-8")
        quadSono = unicode(self.quadSono.currentText()).encode("utf-8")
        orientSono = unicode(self.orientSono.currentText()).encode("utf-8")
        formSono = unicode(self.formSono.currentText()).encode("utf-8")
        echoSono = unicode(self.echoSono.currentText()).encode("utf-8")
        postSono = unicode(self.postSono.currentText()).encode("utf-8")
        randSono = unicode(self.randSono.currentText()).encode("utf-8")
        sizeSono = unicode(self.sizeSono.text()).encode("utf-8")
        #verkalkung
        kalkSono = int(self.kalkSono.isChecked())
        lokalSono = unicode(self.lokalSono.currentText()).encode("utf-8")
        #weiteres
        archSono = int(self.archSono.isChecked())
        intraSono = int(self.intraSono.isChecked())
        asymmSono = unicode(self.asymmSono.currentText()).encode("utf-8")
        zeichSono = unicode(self.zeichSono.currentText()).encode("utf-8")
        elastSono = unicode(self.elastSono.currentText()).encode("utf-8")
        #Begleitmerkmale
        archBSono = int(self.archBSono.isChecked())
        gangSono = int(self.gangSono.isChecked())
        hautSono = int(self.hautSono.isChecked())
        cutisSono = int(self.cutisSono.isChecked())
        cutisverSono = int(self.cutisverSono.isChecked())
        odemSono = int(self.odemSono.isChecked())
        #Spezial
        einfacheZysteSono = int(self.einfacheZysteSono.isChecked())
        mikrozysteSono = int(self.mikrozysteSono.isChecked())
        komplizierteZysteSono = int(self.komplizierteZysteSono.isChecked())
        herdSSono = int(self.herdSSono.isChecked())
        fremdSono = int(self.fremdSono.isChecked())
        intraSSono = int(self.intraSSono.isChecked())
        axillSono = int(self.axillSono.isChecked())
        postopSono = int(self.postopSono.isChecked())
        fettSono = int(self.fettSono.isChecked())

        #TODO
        #Logic to stop unwanted datainput

        #append to list
        values = []
        values.append(studyid)
        values.append(vname)
        values.append(nname)
        values.append(histo)
        values.append(dateSono)
        values.append(internSono)
        #biradsParen
        values.append(biradLeft)
        values.append(biradRight)
        values.append(parenLeft)
        values.append(parenRight)
        #Befund
        values.append(herdSono)
        values.append(zweitSono)
        values.append(laterSono)
        values.append(quadSono)
        values.append(formSono)
        values.append(orientSono)
        values.append(randSono)
        values.append(echoSono)
        values.append(postSono)
        values.append(sizeSono)
        #Verkalkung
        values.append(kalkSono)
        values.append(lokalSono)
        #Weiteres
        values.append(archSono)
        values.append(intraSono)
        values.append(asymmSono)
        values.append(zeichSono)
        values.append(elastSono)
        #Begleitmerkmale
        values.append(archBSono)
        values.append(gangSono)
        values.append(hautSono)
        values.append(cutisSono)
        values.append(cutisverSono)
        values.append(odemSono)
        #Spezial
        values.append(einfacheZysteSono)
        values.append(mikrozysteSono)
        values.append(komplizierteZysteSono)
        values.append(herdSSono)
        values.append(fremdSono)
        values.append(intraSSono)
        values.append(axillSono)
        values.append(postopSono)
        values.append(fettSono)

        #append to datafile in /data/DataMammo #TODO ORDERCheck
        df = pd.read_excel(fnameSono,index=False,encoding='utf-8')
        headers = df.columns.tolist()
        dictTemp = OrderedDict(zip(headers,values))
        newRow = pd.DataFrame([dictTemp])
        df = df.append(newRow)
        df.to_excel(fnameSono, index=False, header = 1,encoding='utf-8')

    def saveMamData(self):
        #check if file exists
        if self.mamCheck() == 0:
            self.createMam()

        #get the Data
        #name etc
        vname = unicode(self.vname.text()).encode("utf-8")
        nname = unicode(self.nname.text()).encode("utf-8")
        studyid = unicode(self.studyid.text()).encode("utf-8")
        histo = unicode(self.histo.currentText()).encode("utf-8")
        #date intern
        dateMam = unicode(self.dateMam.date().toString("dd.MM.yy"))
        internMam = int(self.internMam.isChecked())
        #birads and Parent
        biradLeft = unicode(self.biradLeftMam.currentText()).encode("utf-8") #if none is selected bottom index =0
        biradRight = unicode(self.biradRightMam.currentText()).encode("utf-8")
        parenLeft = unicode(self.parenLeftMam.currentText()).encode("utf-8")
        parenRight = unicode(self.parenRightMam.currentText()).encode("utf-8")
        #Befund
        herdMam = int(self.herdMam.isChecked())
        zweitMam = int(self.zweitMam.isChecked())
        laterMam = unicode(self.laterMam.currentText()).encode("utf-8")
        quadMam = unicode(self.quadMam.currentText()).encode("utf-8")
        depthMam = unicode(self.depthMam.currentText()).encode("utf-8")
        formMam = unicode(self.formMam.currentText()).encode("utf-8")
        randMam = unicode(self.randMam.currentText()).encode("utf-8")
        dichteMam = unicode(self.dichteMam.currentText()).encode("utf-8")
        sizeMam =  unicode(self.sizeMam.text()).encode("utf-8")
        #Verkalkung
        kalkMam = int(self.kalkMam.isChecked())
        typischMam = int(self.typischMam.isChecked())
        typischSusMam = unicode(self.typischSusMam.currentText()).encode("utf-8")
        verteilungMam = unicode(self.verteilungMam.currentText()).encode("utf-8")
        #Weiteres
        archMam = int(self.archMam.isChecked())
        intraMam = int(self.intraMam.isChecked())
        asymmMam = unicode(self.asymmMam.currentText()).encode("utf-8")
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

    def handleButton(self):
        try:
            subprocess.Popen([sys.executable, FILEPATH])
        except OSError as exception:
            print('ERROR: could not restart aplication:')
            print('  %s' % str(exception))
        else:
            QtGui.qApp.quit()


    #TODO Refactor: check since it is now obsolete

    def mamCheck(self):
        #the //Data folder needs to be existent
        return os.path.isfile(fnameMammo)

    def sonoCheck(self):
        #the //Data folder needs to be existent
        return os.path.isfile(fnameSono)

    def mrtCheck(self):
        #the //Data folder needs to be existent
        return os.path.isfile(fnameMrt)

    def createMam(self):
        directory = os.path.dirname(fnameMammo)
        if not os.path.exists(directory):
            os.makedirs(directory)
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'mammo.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMama.xlxs"
        df.to_excel(fnameMammo, index=False)

    def createSono(self):
        directory = os.path.dirname(fnameSono)
        if not os.path.exists(directory):
            os.makedirs(directory)
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'sono.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMama.xlxs"
        df.to_excel(fnameSono, index=False)

    def createMrt(self):
        directory = os.path.dirname(fnameMrt)
        if not os.path.exists(directory):
            os.makedirs(directory)
        #get template file
        template = os.path.join(os.path.curdir,'Templates', 'mrt.xlsx')
        df = pd.read_excel(template)
        #save under "data/DataMrt.xlxs"
        df.to_excel(fnameMrt, index=False)


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
