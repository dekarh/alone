# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'alone_win.ui'
#
# Created by: PyQt5 UI code generator 5.10.1
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1269, 258)
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.frmMenu = QtWidgets.QFrame(Form)
        self.frmMenu.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frmMenu.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frmMenu.setObjectName("frmMenu")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.frmMenu)
        self.horizontalLayout_10.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.clbRefresh = QtWidgets.QCommandLinkButton(self.frmMenu)
        self.clbRefresh.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbRefresh.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("refresh.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbRefresh.setIcon(icon)
        self.clbRefresh.setDescription("")
        self.clbRefresh.setObjectName("clbRefresh")
        self.horizontalLayout_10.addWidget(self.clbRefresh)
        self.calBirtday = QtWidgets.QDateEdit(self.frmMenu)
        self.calBirtday.setAccessibleDescription("")
        self.calBirtday.setMinimumDate(QtCore.QDate(1900, 1, 1))
        self.calBirtday.setCalendarPopup(True)
        self.calBirtday.setDate(QtCore.QDate(1901, 1, 1))
        self.calBirtday.setObjectName("calBirtday")
        self.horizontalLayout_10.addWidget(self.calBirtday)
        self.pbSortF = QtWidgets.QPushButton(self.frmMenu)
        self.pbSortF.setObjectName("pbSortF")
        self.horizontalLayout_10.addWidget(self.pbSortF)
        self.pbSortIO = QtWidgets.QPushButton(self.frmMenu)
        self.pbSortIO.setObjectName("pbSortIO")
        self.horizontalLayout_10.addWidget(self.pbSortIO)
        self.pbSortO = QtWidgets.QPushButton(self.frmMenu)
        self.pbSortO.setObjectName("pbSortO")
        self.horizontalLayout_10.addWidget(self.pbSortO)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_10.addItem(spacerItem)
        self.clbNotFindedXLSX = QtWidgets.QCommandLinkButton(self.frmMenu)
        self.clbNotFindedXLSX.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbNotFindedXLSX.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("saveXLSX.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbNotFindedXLSX.setIcon(icon1)
        self.clbNotFindedXLSX.setDescription("")
        self.clbNotFindedXLSX.setObjectName("clbNotFindedXLSX")
        self.horizontalLayout_10.addWidget(self.clbNotFindedXLSX)
        self.clbRefreshReport = QtWidgets.QCommandLinkButton(self.frmMenu)
        self.clbRefreshReport.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbRefreshReport.setText("")
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("report.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbRefreshReport.setIcon(icon2)
        self.clbRefreshReport.setDescription("")
        self.clbRefreshReport.setObjectName("clbRefreshReport")
        self.horizontalLayout_10.addWidget(self.clbRefreshReport)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_10.addItem(spacerItem1)
        self.clbReport2xlsx = QtWidgets.QCommandLinkButton(self.frmMenu)
        self.clbReport2xlsx.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbReport2xlsx.setText("")
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("saveRED.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbReport2xlsx.setIcon(icon3)
        self.clbReport2xlsx.setDescription("")
        self.clbReport2xlsx.setObjectName("clbReport2xlsx")
        self.horizontalLayout_10.addWidget(self.clbReport2xlsx)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_10.addItem(spacerItem2)
        self.leFile = QtWidgets.QLineEdit(self.frmMenu)
        self.leFile.setObjectName("leFile")
        self.horizontalLayout_10.addWidget(self.leFile)
        self.cbFolder = QtWidgets.QComboBox(self.frmMenu)
        self.cbFolder.setObjectName("cbFolder")
        self.horizontalLayout_10.addWidget(self.cbFolder)
        self.clbSave = QtWidgets.QCommandLinkButton(self.frmMenu)
        self.clbSave.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbSave.setText("")
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("ok.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbSave.setIcon(icon4)
        self.clbSave.setDescription("")
        self.clbSave.setObjectName("clbSave")
        self.horizontalLayout_10.addWidget(self.clbSave)
        self.verticalLayout_6.addWidget(self.frmMenu)
        self.twRez = QtWidgets.QTableWidget(Form)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.twRez.sizePolicy().hasHeightForWidth())
        self.twRez.setSizePolicy(sizePolicy)
        self.twRez.setObjectName("twRez")
        self.twRez.setColumnCount(0)
        self.twRez.setRowCount(0)
        self.verticalLayout_6.addWidget(self.twRez)
        self.frmSNILS = QtWidgets.QFrame(Form)
        self.frmSNILS.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.frmSNILS.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frmSNILS.setObjectName("frmSNILS")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.frmSNILS)
        self.horizontalLayout.setContentsMargins(-1, 0, 9, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.clbLoadBLUE = QtWidgets.QCommandLinkButton(self.frmSNILS)
        self.clbLoadBLUE.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbLoadBLUE.setText("")
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap("saveBLUE.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbLoadBLUE.setIcon(icon5)
        self.clbLoadBLUE.setDescription("")
        self.clbLoadBLUE.setObjectName("clbLoadBLUE")
        self.horizontalLayout.addWidget(self.clbLoadBLUE)
        self.lbSNILS = QtWidgets.QLabel(self.frmSNILS)
        self.lbSNILS.setMaximumSize(QtCore.QSize(360, 16777215))
        self.lbSNILS.setObjectName("lbSNILS")
        self.horizontalLayout.addWidget(self.lbSNILS)
        self.leSNILS = QtWidgets.QLineEdit(self.frmSNILS)
        self.leSNILS.setMaximumSize(QtCore.QSize(130, 16777215))
        self.leSNILS.setText("")
        self.leSNILS.setObjectName("leSNILS")
        self.horizontalLayout.addWidget(self.leSNILS)
        self.clbSNILS = QtWidgets.QCommandLinkButton(self.frmSNILS)
        self.clbSNILS.setMaximumSize(QtCore.QSize(33, 16777215))
        self.clbSNILS.setText("")
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap("right.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.clbSNILS.setIcon(icon6)
        self.clbSNILS.setDescription("")
        self.clbSNILS.setObjectName("clbSNILS")
        self.horizontalLayout.addWidget(self.clbSNILS)
        self.progressBar = QtWidgets.QProgressBar(self.frmSNILS)
        self.progressBar.setEnabled(True)
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.horizontalLayout.addWidget(self.progressBar)
        self.lbDateTime = QtWidgets.QLabel(self.frmSNILS)
        self.lbDateTime.setText("")
        self.lbDateTime.setObjectName("lbDateTime")
        self.horizontalLayout.addWidget(self.lbDateTime)
        self.verticalLayout_6.addWidget(self.frmSNILS)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.calBirtday.setDisplayFormat(_translate("Form", "dd.MM.yyyy"))
        self.pbSortF.setText(_translate("Form", "Фамилия"))
        self.pbSortIO.setText(_translate("Form", "Имя-Отчество"))
        self.pbSortO.setText(_translate("Form", "Отчество"))
        self.lbSNILS.setText(_translate("Form", "СНИЛС для поиска даты звонка в НПФ Социум:"))

