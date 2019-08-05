# -*- coding: utf-8 -*-
__author__ = 'Denis'

import sys

from PyQt5.QtWidgets import (QApplication, QWidget)
from PyQt5.QtGui import QIcon

from alone_slots import MainWindowSlots

class MainWindow(MainWindowSlots):

    # При инициализации класса нам необходимо выполнить некоторые операции
    def __init__(self, form):
        # Сконфигурировать интерфейс методом из базового класса Ui_Form
        self.setupUi(form)
        # Подключить созданные нами слоты к виджетам
        self.connect_slots()

    # Подключаем слоты к виджетам (для каждого действия, которое надо обработать - свой слот)
    def connect_slots(self):
        self.clbRefresh.clicked.connect(self.click_clbRefresh)
        self.clbRefreshReport.clicked.connect(self.click_clbRefreshReport)
        self.clbReport2xlsx.clicked.connect(self.click_clbReport2xlsx)
        self.clbSave.clicked.connect(self.click_clbSave)
        self.pbSortF.clicked.connect(self.click_pbSortF)
        self.pbSortIO.clicked.connect(self.click_pbSortIO)
        self.pbSortO.clicked.connect(self.click_pbSortO)
        self.leFile.textChanged[str].connect(self.leFile_changed)
        self.cbFolder.activated[str].connect(self.click_cbFolder)
        self.twRez.clicked.connect(self.click_twRez)
        return None

if __name__ == '__main__':
    # Создаём экземпляр приложения
    app = QApplication(sys.argv)
    # Создаём базовое окно, в котором будет отображаться наш UI
    window = QWidget()
    window.setWindowIcon(QIcon('alone.png'))
    # Создаём экземпляр нашего UI
    ui = MainWindow(window)
    # Отображаем окно
    window.show()
    # Обрабатываем нажатие на кнопку окна "Закрыть"
    sys.exit(app.exec_())




