# -*- coding: utf-8 -*-
# для поиска по базе адресов нужно стартовать сервисы sphinx и fias

from subprocess import Popen, PIPE
import os
import sys
import re
import string
import bz2
from string import digits
from random import random
from dateutil.parser import parse
from collections import OrderedDict

from datetime import datetime, timedelta, time, date
#from time import time
import pytz
utc=pytz.UTC

import openpyxl
from openpyxl import Workbook
import requests, json


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QDate, QDateTime, QSize, Qt, QByteArray, QTimer, QUrl, QThread
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox, QMainWindow, QWidget, QFrame, QFileDialog, QComboBox

from mysql.connector import MySQLConnection

from alone_win import Ui_Form

# import NormalizeFields as norm
from lib import read_config, l, s, fine_phone, format_phone

class MainWindowSlots(Ui_Form):   # Определяем функции, которые будем вызывать в слотах

    def setupUi(self, form):
        Ui_Form.setupUi(self,form)
        self.client_id = None
        self.hasFileFolder = False
        self.dbconfig_crm = read_config(filename='alone.ini', section='crm')
        self.dbconfig_alone = read_config(filename='alone.ini', section='alone')
        self.alone_files = {}
        with open("all_files.txt", "rt") as file_all:
            for i, line in enumerate(file_all):
                if i > 1:
                    if len(line.split('/')) > 2 and line.find('search') == -1:
                        file_name = line.split('.wav')[0].split('/')[2].lower()
                        path_name = line.split('./recup_dir.')[1].split('/')[0].replace('/n','')
                        if self.alone_files.get(file_name, None):
                            self.alone_files[file_name].append(path_name)
                        else:
                            self.alone_files[file_name] = [path_name]
        self.twRezkeyPressEventMain = self.twRez.keyPressEvent
        self.twRez.keyPressEvent = self.twRezkeyPressEvent
        self.clbSave.setEnabled(False)
        self.contracts = {None:None}
        self.clbReport2xlsx.setEnabled(False)
        return

    def twRezkeyPressEvent(self,e):
        self.twRezkeyPressEventMain(e)
        if e.key() == Qt.Key_Down or e.key() == Qt.Key_Up:
            self.click_twRez(index=self.twRez.model().index(self.twRez.currentRow(), 0))

    def click_twRez(self, index=None): # Сделать кнопку Сохранить активной если есть файл, папка и выбран договор
        self.client_id =  self.client_ids[index.row()]
        if self.hasFileFolder and self.client_id: # Сделать кнопку Сохранить активной если есть файл, папка и выбран договор
            self.clbSave.setEnabled(True)
        else:
            self.clbSave.setEnabled(False)

    def click_cbFolder(self):
        if len(self.cbFolder.currentText()):
            self.hasFileFolder = True
        else:
            self.hasFileFolder = False
        if self.hasFileFolder and self.client_id:
            self.clbSave.setEnabled(True)
        else:
            self.clbSave.setEnabled(False)

    def leFile_changed(self):
        self.hasFileFolder = False
        self.clbSave.setEnabled(False)
        if self.alone_files.get(self.leFile.text(), None):
            self.cbFolder.clear()
            self.cbFolder.addItems(self.alone_files[self.leFile.text()])

    def click_clbRefresh(self):
        if self.calBirtday.dateTime().toPyDateTime().date() > date(1930,1,1) and \
                                self.calBirtday.dateTime().toPyDateTime() < datetime.now():
            self.leFile.setText('')
            self.hasFileFolder = False
            self.client_id = None
            self.clbSave.setEnabled(False)
            dbconn = MySQLConnection(**self.dbconfig_crm)
            cursor = dbconn.cursor()
            sql = 'SELECT cl.p_surname,cl.p_name,cl.p_lastname,cl.p_service_address,cl.d_service_address,' \
                  'ca.client_phone,ca.call_comment,ca.inserted_date,cl.client_id FROM saturn_crm.clients AS cl ' \
                  'LEFT JOIN saturn_crm.contracts AS co ON co.client_id = cl.client_id ' \
                  'LEFT JOIN saturn_crm.callcenter AS ca ON ca.contract_id = co.id ' \
                  'WHERE cl.b_date = %s'
            cursor.execute(sql, (self.calBirtday.dateTime().toPyDateTime(),))
            rows = cursor.fetchall()
            dogovors = {}
            for row in rows:
                client_id = row[8]
                if dogovors.get(client_id, None):
                    if row[7].date() not in dogovors[client_id]['Даты']:
                        dogovors[client_id]['Даты'] = dogovors[client_id]['Даты'] + [row[7].date()]
                else:
                    dogovor = {}
                    dogovor['client_id'] = client_id
                    dogovor['Фамилия'] = row[0]
                    dogovor['Имя'] = row[1]
                    dogovor['Отчество'] = row[2]
                    dogovor['Регистрация'] = row[3]
                    dogovor['Проживание'] = row[4]
                    dogovor['Телефон'] = row[5]
                    dogovor['Коментарий'] = row[6]
                    if row[7]:
                        dogovor['Даты'] = [row[7].date()]
                    else:
                        dogovor['Даты'] = [None]
                    dogovors[client_id] = dogovor
            self.contracts = {}
            contracts4order = {}
            for client_id in dogovors:
                if dogovors[client_id]['Даты'] != [None]:
                    self.contracts[client_id] = dogovors[client_id]
                    contracts4order[client_id] = dogovors[client_id]['Фамилия']
            keys = ['Фамилия', 'Имя', 'Отчество', 'Регистрация', 'Проживание', 'Телефон', 'Коментарий', 'Даты']
            self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
            self.twRez.setRowCount(len(contracts4order))  # Кол-во строк из таблицы
            contracts_ordered = OrderedDict(sorted(contracts4order.items(), key=lambda t: t[1]))
            self.client_ids = []
            for j, client_id in enumerate(contracts_ordered):
                self.client_ids.append(client_id)
                for k, key in enumerate(keys):
                    if key == 'Даты':
                        if self.contracts[client_id].get('Даты', False):
                            all_dates = ';'.join([data.strftime('%d.%m.%y') for data in self.contracts[client_id][key]])
                            self.twRez.setItem(j, k, QTableWidgetItem(all_dates))
                    else:
                        self.twRez.setItem(j, k, QTableWidgetItem(str(self.contracts[client_id][key])))
            # Устанавливаем заголовки таблицы
            self.twRez.setHorizontalHeaderLabels(list(keys))
            # Устанавливаем выравнивание на заголовки
            self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
            # делаем ресайз колонок по содержимому
            self.twRez.horizontalHeader().resizeSection(0, 150)
            self.twRez.horizontalHeader().resizeSection(1, 100)
            self.twRez.horizontalHeader().resizeSection(2, 150)
            self.twRez.horizontalHeader().resizeSection(3, 250)
            self.twRez.horizontalHeader().resizeSection(4, 250)
            self.twRez.horizontalHeader().resizeSection(5, 100)
            self.twRez.horizontalHeader().resizeSection(6, 100)
            self.twRez.horizontalHeader().resizeSection(7, 100)
            return

    def click_clbSave(self):
        self.clbReport2xlsx.setEnabled(False)
        dbconn = MySQLConnection(**self.dbconfig_alone)
        cursor = dbconn.cursor()
        sql = 'SELECT * FROM alone_connect WHERE path = %s AND file = %s AND client_id = %s'
        cursor.execute(sql, (self.cbFolder.currentText(), self.leFile.text(), self.client_id))
        rows = cursor.fetchall()
        if len(rows) == 0:
            cursor = dbconn.cursor()
            cursor.execute('INSERT INTO alone_connect (path, file, client_id) VALUES(%s, %s, %s)',
                           (self.cbFolder.currentText(), self.leFile.text(), self.client_id))
            dbconn.commit()
        dbconn.close()

    def click_pbSortF(self):
        contracts4order = {}
        for client_id in self.contracts:
            if self.contracts[client_id]['Даты'] != [None]:
                contracts4order[client_id] = self.contracts[client_id]['Фамилия']
        keys = ['Фамилия', 'Имя', 'Отчество', 'Регистрация', 'Проживание', 'Телефон', 'Коментарий', 'Даты']
        self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
        self.twRez.setRowCount(len(contracts4order))  # Кол-во строк из таблицы
        contracts_ordered = OrderedDict(sorted(contracts4order.items(), key=lambda t: t[1]))
        self.client_ids = []
        for j, client_id in enumerate(contracts_ordered):
            self.client_ids.append(client_id)
            for k, key in enumerate(keys):
                if key == 'Даты':
                    if self.contracts[client_id].get('Даты', False):
                        all_dates = ';'.join([data.strftime('%d.%m.%y') for data in self.contracts[client_id][key]])
                        self.twRez.setItem(j, k, QTableWidgetItem(all_dates))
                else:
                    self.twRez.setItem(j, k, QTableWidgetItem(str(self.contracts[client_id][key])))
        # Устанавливаем заголовки таблицы
        self.twRez.setHorizontalHeaderLabels(list(keys))
        # Устанавливаем выравнивание на заголовки
        self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twRez.horizontalHeader().resizeSection(0, 150)
        self.twRez.horizontalHeader().resizeSection(1, 100)
        self.twRez.horizontalHeader().resizeSection(2, 150)
        self.twRez.horizontalHeader().resizeSection(3, 250)
        self.twRez.horizontalHeader().resizeSection(4, 250)
        self.twRez.horizontalHeader().resizeSection(5, 100)
        self.twRez.horizontalHeader().resizeSection(6, 100)
        self.twRez.horizontalHeader().resizeSection(7, 100)

    def click_pbSortO(self):
        contracts4order = {}
        contracts4orderNone = {}
        for client_id in self.contracts:
            if self.contracts[client_id]['Даты'] != [None]:
                if self.contracts[client_id]['Отчество']:
                    contracts4order[client_id] = self.contracts[client_id]['Отчество']
                else:
                    contracts4orderNone[client_id] = self.contracts[client_id]['Отчество']
        keys = ['Фамилия', 'Имя', 'Отчество', 'Регистрация', 'Проживание', 'Телефон', 'Коментарий', 'Даты']
        self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
        self.twRez.setRowCount(len(contracts4order))  # Кол-во строк из таблицы
        contracts_ordered = OrderedDict(sorted(contracts4order.items(), key=lambda t: t[1]))
        for client_id in contracts4orderNone:
            contracts_ordered[client_id] = contracts4orderNone[client_id]
        self.client_ids = []
        for j, client_id in enumerate(contracts_ordered):
            self.client_ids.append(client_id)
            for k, key in enumerate(keys):
                if key == 'Даты':
                    if self.contracts[client_id].get('Даты', False):
                        all_dates = ';'.join([data.strftime('%d.%m.%y') for data in self.contracts[client_id][key]])
                        self.twRez.setItem(j, k, QTableWidgetItem(all_dates))
                else:
                    self.twRez.setItem(j, k, QTableWidgetItem(str(self.contracts[client_id][key])))
        # Устанавливаем заголовки таблицы
        self.twRez.setHorizontalHeaderLabels(list(keys))
        # Устанавливаем выравнивание на заголовки
        self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twRez.horizontalHeader().resizeSection(0, 150)
        self.twRez.horizontalHeader().resizeSection(1, 100)
        self.twRez.horizontalHeader().resizeSection(2, 150)
        self.twRez.horizontalHeader().resizeSection(3, 250)
        self.twRez.horizontalHeader().resizeSection(4, 250)
        self.twRez.horizontalHeader().resizeSection(5, 100)
        self.twRez.horizontalHeader().resizeSection(6, 100)
        self.twRez.horizontalHeader().resizeSection(7, 100)

    def click_pbSortIO(self):
        contracts4order = {}
        for client_id in self.contracts:
            if self.contracts[client_id]['Даты'] != [None]:
                contracts4order[client_id] = self.contracts[client_id]['Имя']
        keys = ['Фамилия', 'Имя', 'Отчество', 'Регистрация', 'Проживание', 'Телефон', 'Коментарий', 'Даты']
        self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
        self.twRez.setRowCount(len(contracts4order))  # Кол-во строк из таблицы
        contracts_ordered = OrderedDict(sorted(contracts4order.items(), key=lambda t: t[1]))
        self.client_ids = []
        for j, client_id in enumerate(contracts_ordered):
            self.client_ids.append(client_id)
            for k, key in enumerate(keys):
                if key == 'Даты':
                    if self.contracts[client_id].get('Даты', False):
                        all_dates = ';'.join([data.strftime('%d.%m.%y') for data in self.contracts[client_id][key]])
                        self.twRez.setItem(j, k, QTableWidgetItem(all_dates))
                else:
                    self.twRez.setItem(j, k, QTableWidgetItem(str(self.contracts[client_id][key])))
        # Устанавливаем заголовки таблицы
        self.twRez.setHorizontalHeaderLabels(list(keys))
        # Устанавливаем выравнивание на заголовки
        self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twRez.horizontalHeader().resizeSection(0, 150)
        self.twRez.horizontalHeader().resizeSection(1, 100)
        self.twRez.horizontalHeader().resizeSection(2, 150)
        self.twRez.horizontalHeader().resizeSection(3, 250)
        self.twRez.horizontalHeader().resizeSection(4, 250)
        self.twRez.horizontalHeader().resizeSection(5, 100)
        self.twRez.horizontalHeader().resizeSection(6, 100)
        self.twRez.horizontalHeader().resizeSection(7, 100)

    def click_clbRefreshReport(self):
        dbconn = MySQLConnection(**self.dbconfig_alone)
        cursor = dbconn.cursor()
        cursor.execute('SELECT client_id, path, file FROM alone_connect')
        rows = cursor.fetchall()
        temp_ids = []
        report_client_ids = {} # Даже индекс из папки+файл тоже может повторяться ((( Добавляем номер дубля (i) вначале
        for row in rows:
            temp_ids.append(row[0])
            for i in range (0,9):
                if report_client_ids.get(str(i) + '{0:04d}'.format(int(row[1])) + row[2], None):
                    pass
                else:
                    report_client_ids[str(i) + '{0:04d}'.format(int(row[1])) + row[2]] = row[0]
                    break
        uniq_client_ids = list(set(temp_ids)) # Убираем повторы из массива idшников чтобы
                                              # запросить внутренности нужных договоров
        dbconn = MySQLConnection(**self.dbconfig_crm)
        cursor = dbconn.cursor()
        sql = 'SELECT cl.p_surname,cl.p_name,cl.p_lastname,cl.p_service_address,cl.d_service_address,' \
              'ca.client_phone,ca.call_comment,ca.inserted_date,cl.client_id,cl.b_date FROM saturn_crm.clients AS cl ' \
              'LEFT JOIN saturn_crm.contracts AS co ON co.client_id = cl.client_id ' \
              'LEFT JOIN saturn_crm.callcenter AS ca ON ca.contract_id = co.id ' \
              'WHERE cl.client_id in ({c})'.format(c=', '.join(['%s'] * len(uniq_client_ids)))
        cursor.execute(sql, tuple(uniq_client_ids))
        rows = cursor.fetchall()
        dogovors = {}
        for row in rows:
            client_id = row[8]
            if dogovors.get(client_id, None):
                if row[7].date() not in dogovors[client_id]['Даты']:
                    dogovors[client_id]['Даты'] = dogovors[client_id]['Даты'] + [row[7].date()]
            else:
                dogovor = {}
                dogovor['client_id'] = client_id
                dogovor['Фамилия'] = row[0]
                dogovor['Имя'] = row[1]
                dogovor['Отчество'] = row[2]
                dogovor['Регистрация'] = row[3]
                dogovor['Проживание'] = row[4]
                dogovor['Телефон'] = row[5]
                dogovor['Коментарий'] = row[6]
                if row[7]:
                    dogovor['Даты'] = [row[7].date()]
                else:
                    dogovor['Даты'] = [None]
                dogovor['ДеньРождения'] = row[9]
                dogovors[client_id] = dogovor
        report = {}
        for report_client_id in report_client_ids:
            path = int(report_client_id[1:5])  #file = report_client_id[5:]
            client_id = report_client_ids[report_client_id]
            dates = dogovors[client_id]['Даты']
            if report.get(path, None):
                # есть такая папка
                if report[path].get(client_id, None):
                    # есть такая папка и такой client_id
                    for data in report[path][client_id]:
                        if data not in dates:
                            dates = dates + [data]
                    report[path][client_id] = dates
                else:
                    # есть такая папка и нет такого client_id !!! первая дата - телефон
                    report[path][client_id] = [dogovors[client_id]['Телефон']] + dates
            else:
                # нет такой папки !!! первая дата - телефон
                report[path] = {client_id: [dogovors[client_id]['Телефон']] + dates}
        # перестраиваем с client_id на телефоны
        report2phones = {}
        for path in report:
            report2phones[path] = {}
            for client_id in report[path]:
                phone = report[path][client_id][0]
                if report2phones[path].get(phone, None):
                    dates = report2phones[path][phone]
                    # есть такая папка и такой телефон
                    for i, data in enumerate(report[path][client_id]):
                        if i:
                            if data not in dates:
                                dates = dates + [data]
                    report2phones[path][phone] = dates
                else:
                    # есть такая папка и нет такого телефона
                    report2phones[path][phone] = report[path][client_id][1:]
        # анализируем отчет
        report_rez = {}
        for path in report2phones:
            dates = {}
            for phone in report2phones[path]:
                for data in report2phones[path][phone]:
                    if dates.get(data, None):
                        # есть такая дата
                        dates[data] += 1
                    else:
                        dates[data] = 1
            dates_ordered = OrderedDict(sorted(dates.items(), key=lambda t: t[1], reverse=True))
            for data in dates_ordered:
                if len(report2phones[path]) > 1 and dates_ordered[data] >= len(report2phones[path]):
                    report_rez[path] = datetime.combine(data,time(0,0,0,0)).strftime('%d.%m.%y')
                elif len(report2phones[path]) > 1:
                    report_rez[path] = 'Мульти'
                else:
                    report_rez[path] = 'Начато'
                break

        self.clbReport2xlsx.setEnabled(True)

    def click_clbReport2xlsx(self):
#        wb_log = openpyxl.Workbook(write_only=True)
#        ws_log = wb_log.create_sheet('Лог')
#        ws_log.append([datetime.now().strftime("%H:%M:%S"), ' Начинаем'])
#        wb_log.save('1.xlsx')
        return

