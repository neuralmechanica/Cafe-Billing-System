from typing import Text
from PyQt5 import QtWidgets
import pandas as pd
import datetime
from PyQt5.QtCore import *
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QApplication
import sys
import win32ui
import win32print
import win32con
from os import path
from PyQt5.uic import loadUiType
from xlrd import sheet
from_class,_= loadUiType(path.join(path.dirname('__file__'), "main.ui"))


class Main(QMainWindow, from_class):
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.receipt1 = ''
        self.Total_Cost = 0
        self.Handel_Buttons()
        self.reset1()
        self.df = pd.read_excel("Item.xlsx")        





    def Handel_Buttons(self):
        self.reset.clicked.connect(self.reset1)
        self.gr.clicked.connect(self.daam)
        self.print.clicked.connect(self.start_print)


    def reset1(self):
        self.roll_parata.setText('0')
        self.samosa.setText('0')
        self.Vburger.setText('0')
        self.Cburger.setText('0')
        self.Eburger.setText('0')
        self.sandwich.setText('0')
        self.Gtea.setText('0')
        self.Mtea.setText('0')
        self.coffee.setText('0')
        self.Hmilk.setText('0')
        self.Bcoffee.setText('0')
        self.Btea.setText('0')
        self.orng.setText('0')
        self.Jmango.setText('0')
        self.apple.setText('0')
        self.Papple.setText('0')
        self.banana.setText('0')
        self.Smango.setText('0')
        self.almond.setText('0')
        self.sb.setText('0')
        self.cost.setText('0')
        dateNtime = datetime.datetime.now()
        self.receipt1 = str('\t\t\t             %s/%s/%s  %s:%s:%s' %(dateNtime.day, dateNtime.month, dateNtime.year, dateNtime.hour, dateNtime.minute, dateNtime.second ))
        self.receipt.setText(self.receipt1)
        self.Total_Cost = 0

    def daam(self):
        rp = self.roll_parata.toPlainText()
        sa = self.samosa.toPlainText()
        vb = self.Vburger.toPlainText()
        cb = self.Cburger.toPlainText()
        eb = self.Eburger.toPlainText()
        san = self.sandwich.toPlainText()
        gt = self.Gtea.toPlainText()
        mt = self.Mtea.toPlainText()
        co = self.coffee.toPlainText()
        hm = self.Hmilk.toPlainText()
        bco = self.Bcoffee.toPlainText()
        bt = self.Btea.toPlainText()
        orn = self.orng.toPlainText()
        jm = self.Jmango.toPlainText()
        app = self.apple.toPlainText()
        papp = self.Papple.toPlainText()
        ba = self.banana.toPlainText()
        sm = self.Smango.toPlainText()
        al = self.almond.toPlainText()
        sb = self.sb.toPlainText()
        rp = int(rp)
        if rp >= 1:
            price6 = int(self.df.values[0, 1])
            price = rp*price6
            self.receipt1 = '%s \n Roll Parata  X  %s  =  %s'%(self.receipt1, rp, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        sa = int(sa)
        if  sa >= 1:
            price6 = int(self.df.values[1, 1])
            price = sa*price6
            self.receipt1 = '%s \n Samosa  X  %s  =  %s'%(self.receipt1, sa, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        vb = int(vb)
        if  vb >= 1:
            price6 = int(self.df.values[2, 1])
            price = vb*price6
            self.receipt1 = '%s \n Veg Burger  X  %s  =  %s'%(self.receipt1, vb, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        cb = int(cb)
        if  cb >= 1:
            price6 = int(self.df.values[3, 1])
            price = cb*price6
            self.receipt1 = '%s \n Chiken Burger  X  %s  =  %s'%(self.receipt1, cb, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        eb = int(eb)
        if  eb >= 1:
            price6 = int(self.df.values[4, 1])
            price = eb*price6
            self.receipt1 = '%s \n Egg Burger  X  %s  =  %s'%(self.receipt1, eb, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        san = int(san)
        if  san >= 1:
            price6 = int(self.df.values[5, 1])
            price = san*price6
            self.receipt1 = '%s \n Sandwich  X  %s  =  %s'%(self.receipt1, san, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        gt = int(gt)
        if  gt >= 1:
            price6 = int(self.df.values[6, 1])
            price = gt*price6
            self.receipt1 = '%s \n Green Tea X  %s  =  %s'%(self.receipt1, gt, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        mt = int(mt)
        if  mt >= 1:
            price6 = int(self.df.values[7, 1])
            price = mt*price6
            self.receipt1 = '%s \n Milk Tea  X  %s  =  %s'%(self.receipt1, mt, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        co = int(co)
        if  co >= 1:
            price6 = int(self.df.values[8, 1])
            price = co*price6
            self.receipt1 = '%s \n Coffee  X  %s  =  %s'%(self.receipt1, co, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        hm = int(hm)
        if  hm >= 1:
            price6 = int(self.df.values[9, 1])
            price = hm*price6
            self.receipt1 = '%s \n Hot Milk  X  %s  =  %s'%(self.receipt1, hm, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        bco = int(bco)
        if  bco >= 1:
            price6 = int(self.df.values[10, 1])
            price = bco*price6
            self.receipt1 = '%s \n Black Coffee  X  %s  =  %s'%(self.receipt1, bco, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        bt = int(bt)
        if  bt >= 1:
            price6 = int(self.df.values[11, 1])
            price = bt*price6
            self.receipt1 = '%s \n Black Tea  X  %s  =  %s'%(self.receipt1, bt, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        orn = int(orn)
        if  orn >= 1:
            price6 = int(self.df.values[12, 1])
            price = orn*price6
            self.receipt1 = '%s \n Orange J  X  %s  =  %s'%(self.receipt1, orn, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price
        jm = int(jm)
        if  jm >= 1:
            price6 = int(self.df.values[13, 1])
            price = jm*price6
            self.receipt1 = '%s \n Mango J  X  %s  =  %s'%(self.receipt1, jm, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        app = int(app)
        if  app >= 1:
            price6 = int(self.df.values[14, 1])
            price = app*price6
            self.receipt1 = '%s \n Apple J  X  %s  =  %s'%(self.receipt1, app, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        papp = int(papp)
        if  papp >= 1:
            price6 = int(self.df.values[15, 1])
            price = papp*price6
            self.receipt1 = '%s \n P Apple J  X  %s  =  %s'%(self.receipt1, papp, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        ba = int(ba)
        if  ba >= 1:
            price6 = int(self.df.values[16, 1])
            price = ba*price6
            self.receipt1 = '%s \n Banana  Shake X  %s  =  %s'%(self.receipt1, ba, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        sm = int(sm)
        if  sm >= 1:
            price6 = int(self.df.values[17, 1])
            price = sm*price6
            self.receipt1 = '%s \n Mango Shake  X  %s  =  %s'%(self.receipt1, sm, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        al = int(al)
        if  al >= 1:
            price6 = int(self.df.values[18, 1])
            price = al*price6
            self.receipt1 = '%s \n Almond Shake  X  %s  =  %s'%(self.receipt1, al, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price

        sb = int(sb)
        if  sb >= 1:
            price6 = int(self.df.values[19, 1])
            price = sb*price6
            self.receipt1 = '%s \n Strawberry Shake  X  %s  =  %s'%(self.receipt1, sb, price)
            self.receipt.setText(self.receipt1)
            self.Total_Cost = self.Total_Cost + price 

        self.cost.setText(str(self.Total_Cost))
    
    def start_print(self):
        inch = 1440
        pdc = win32ui.CreateDC()
        pdc.CreatePrinterDC(win32print.GetDefaultPrinter())
        pdc.StartDoc('Main Slip')
        pdc.StartPage()
        pdc.SetMapMode(win32con.MM_TWIPS)
        pdc.DrawText(self.receipt1, (0, inch * -1, inch*8, inch*-2), win32con.DT_CENTER)
        pdc.EndPage()
        pdc.EndDoc()
    

def main():
    app=QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()