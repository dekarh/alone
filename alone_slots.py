# -*- coding: utf-8 -*-
# для поиска по базе адресов нужно стартовать сервисы sphinx и fias

from collections import OrderedDict

from datetime import datetime, timedelta, time, date
import openpyxl

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTableWidgetItem

from mysql.connector import MySQLConnection

from alone_win import Ui_Form

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
        self.threads = []
        return

    def click_clbLoadXlsx(self):
        wb_paths = openpyxl.Workbook(write_only=True)
        ws_unknowns = wb_paths.create_sheet('Аудиозаписи из неопределившейся части файлопомойки')
        ws_unknowns.append(['СНИЛС', 'Дата звонка', 'Ф.И.О.', 'день рождения', 'Прописка'])
        ws_paths = wb_paths.create_sheet('Прослушивание по папкам')
        ws_paths.append(['Папка', 'Ф.И.О.', 'день рождения', 'Прописка'])
        wb = openpyxl.load_workbook(filename='нужноАудио.xlsx', read_only=True)
        ws = wb[wb.sheetnames[0]]
        snilses = []
        for i, row in enumerate(ws):
            if i:
                for j, cell in enumerate(row):
                    if j == 1:
                        snilses.append(l(cell.value))
        dbconn = MySQLConnection(**self.dbconfig_crm)
        cursor = dbconn.cursor()
        for snils in snilses:
            sql = 'SELECT cl.client_id, ca.client_phone, ca.inserted_date, ca.exchangeable, ' \
                  'concat_ws(" ", cl.p_surname, cl.p_name, cl.p_lastname), cl.b_date, cl.p_service_address ' \
                  'FROM saturn_crm.callcenter AS ca ' \
                  'LEFT JOIN saturn_crm.contracts AS co ON ca.contract_id = co.id ' \
                  'LEFT JOIN saturn_crm.clients AS cl ON co.client_id = cl.client_id ' \
                  'WHERE cl.number = %s and cl.subdomain_id = 6'
            cursor.execute(sql, (l(snils),))
            rows = cursor.fetchall()
            if len(rows):
                fio = rows[0][4]
                birthday = rows[0][5]
                address = rows[0][6]
                data = datetime(2001, 1, 1, 0, 0)
                has_checked = False
                for row in rows:
                    if row[2] > data:
                        data = row[2]
                    if row[3]:
                        has_checked = True
                        checked_row = row
                if has_checked:
                    data = checked_row[2]
                pathDataDate = {}
                finded = False
                for thread in self.threads:
                    if data > thread['start'] and data < thread['end']:
                        finded = True
                        max_date_delta = timedelta(days=10000)
                        min_date_delta = timedelta(days=10000)
                        for path in thread['pathsDates']:
                            for call_date in thread['pathsDates'][path]:
                                if call_date > data and (call_date - data) < max_date_delta:
                                    max_date_delta = call_date - data
                                    max_path = path
                                if call_date < data and (data - call_date) < min_date_delta:
                                    min_date_delta = data - call_date
                                    min_path = path
                        for path in range(min_path, max_path + 1):
                            if pathDataDate.get(path, None):
                                if pathDataDate[path].get(snils, None):
                                    pathDataDate[path][snils] = [path, fio, birthday, address]
                                else:
                                    pathDataDate[path] = {snils: [path, fio, birthday, address]}
                            else:
                                pathDataDate[path] = {snils: [path, fio, birthday, address]}
                if not finded:
                    ws_unknowns.append([snils, data, row[4], row[5], row[6]])
            else:
                self.lbDateTime.setText('Нет такого СНИЛС в БД')
        pathDataDate_sorted = OrderedDict(sorted(pathDataDate.items(), key=lambda t: t[0]))
        for path in pathDataDate_sorted:
            for snils in pathDataDate[path]:
                ws_paths.append(pathDataDate[path][snils])
        wb_paths.save('нужноАудио-Отчет.xlsx')
        return

    def click_clbSNILS(self):
        self.lbDateTime.setText('')
        if l(self.leSNILS.text()) < 10000 or l(self.leSNILS.text()) > 99999999999:
            self.lbDateTime.setText('')
        else:
            dbconn = MySQLConnection(**self.dbconfig_crm)
            cursor = dbconn.cursor()
            sql = 'SELECT cl.client_id, ca.client_phone, ca.inserted_date, ca.exchangeable ' \
                  'FROM saturn_crm.callcenter AS ca ' \
                  'LEFT JOIN saturn_crm.contracts AS co ON ca.contract_id = co.id ' \
                  'LEFT JOIN saturn_crm.clients AS cl ON co.client_id = cl.client_id ' \
                  'WHERE cl.number = %s and cl.subdomain_id = 6'
            cursor.execute(sql, (l(self.leSNILS.text()),))
            rows = cursor.fetchall()
            if len(rows):
                data = datetime(2001,1,1,0,0)
                has_checked = False
                for row in rows:
                    if row[2] > data:
                        data = row[2]
                    if row[3]:
                        has_checked = True
                        checked_row = row
                if has_checked:
                    data = checked_row[2]
                textDateTime = data.strftime('%d.%m.%Y') + ' прослушать папки'
                finded = False
                for thread in self.threads:
                    if data > thread['start'] and data < thread['end']:
                        finded = True
                        max_date_delta = timedelta(days=10000)
                        min_date_delta = timedelta(days=10000)
                        for path in thread['pathsDates']:
                            for call_date in thread['pathsDates'][path]:
                                if call_date > data and (call_date - data) < max_date_delta:
                                    max_date_delta = call_date - data
                                    max_path = path
                                if call_date < data and (data - call_date) < min_date_delta:
                                    min_date_delta = data - call_date
                                    min_path = path
                        if textDateTime[-16:] == 'прослушать папки':
                            textDateTime += ' c ' + str(min_path) + ' по ' + str(max_path)
                        else:
                            textDateTime += ', c ' + str(min_path) + ' по ' + str(max_path)
                if not finded:
                    textDateTime += ' недостаточно информации для определения'
                self.lbDateTime.setText(textDateTime)

            else:
                self.lbDateTime.setText('Нет такого СНИЛС в БД')
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
            #rows = [('МЕЛЬНИКОВ', 'ВАЛЕНТИН', 'ИВАНОВИЧ', 'Тверская обл, Калязинский р-н, г Калязин, ул 1 Мая, д 5', 'Тверская обл, Калязинский р-н, г Калязин, ул 1 Мая, д 5', 79857795218, 'нпф название не помнит', datetime(2016, 4, 4, 13, 24, 46), '64c46542-fa4e-11e5-9847-5254004b76e6'), ('ХОДЕНЁВА', 'НАТАЛЬЯ', 'СЕРГЕЕВНА', 'Алтайский край, Романовский р-н, село Сидоровка, ул Партизанская, д 11А, кв 2', 'Алтайский край, г Барнаул, ул Антона Петрова, д 176, кв 50', 79237160083, '', datetime(2016, 5, 14, 8, 31, 1), '1272d548-130a-11e6-8b81-5254004b76e6')]
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
        q1 = """
        # Заполняем информацию по прослушиванию оператором
        dbconn = MySQLConnection(**self.dbconfig_alone)
        cursor = dbconn.cursor()
        cursor.execute('SELECT client_id, path, file FROM alone_connect')
        rows = cursor.fetchall()
        #rows = [('39f07f6d-16e7-11e8-86b5-5254004b76e6', '1', 'f658161664'), ('8113962c-16b8-11e8-86b5-5254004b76e6', '2', 'f659636224'), ('42610b96-16ea-11e8-86b5-5254004b76e6', '2', 'f659652608'), ('f51b5fd6-178d-11e8-86b5-5254004b76e6', '2', 'f687063040'), ('8c95a423-bbc5-11e6-b8cb-20cf300dec24', '6', 'f1083621376'), ('820c39ab-178f-11e8-86b5-5254004b76e6', '6', 'f1083621376'), ('3c0652e8-1809-11e8-81ec-5254004b76e6', '7', 'f2880913408'), ('d2d14811-18a0-11e8-81ec-5254004b76e6', '8', 'f3592290304'), ('b77d6b73-04ae-11e7-9f62-5254004b76e6', '9', 'f3592732672'), ('d30b3605-180c-11e8-81ec-5254004b76e6', '9', 'f3592732672'), ('525e4b86-d737-11e6-aa92-20cf300dec24', '9', 'f3712696320'), ('d8fa8330-178f-11e8-86b5-5254004b76e6', '9', 'f3712696320'), ('d3950684-fd16-11e8-8408-000c290cfc84', '25', 'f3909402624'), ('2c80db53-fddc-11e8-8408-000c290cfc84', '25', 'f3899719680'), ('c3b71aed-1da7-11e7-8786-5254004b76e6', '150', 'f2267316224'), ('f2c3a29d-156f-11e8-9039-5254004b76e6', '150', 'f2267807744'), ('d758b226-1808-11e8-81ec-5254004b76e6', '150', 'f2267807744'), ('3830045c-047b-11e9-a9ee-000c290cfc84', '150', 'f2267807744'), ('e4eba1e7-9010-11e7-8989-5254004b76e6', '3', 'f920420352'), ('f8262d02-ec50-11e7-897e-5254004b76e6', '3', 'f920420352')]
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
        #rows = [('ЮРЬЕВ', 'МИХАИЛ', 'АНАТОЛЬЕВИЧ', 'Омская обл, г Омск, Советский округ, ул Химиков, д 16, кв 61', 'Омская обл, г Омск, Советский округ, ул Химиков, д 16, кв 61', 79831198154, '', datetime(2018, 12, 12, 12, 54, 2), '2c80db53-fddc-11e8-8408-000c290cfc84', date(1989, 6, 15))]
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
        self.report_rez = {}
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
                    self.report_rez[path] = datetime.combine(data,time(0,0,0,0)).strftime('%d.%m.%y')
                elif len(report2phones[path]) > 1:
                    self.report_rez[path] = 'МУЛЬТИ'
                else:
                    self.report_rez[path] = 'начато'
                break
        """
        # Добавляем/заменяем информацию по распознанным файлам
        dbconn = MySQLConnection(**self.dbconfig_crm)
        cursor = dbconn.cursor()
        # Распознанные файлы, отсортированные по директориям
        cursor.execute('SELECT r.`path`, c.inserted_date FROM lekarh.alone_remont AS r '
                       'LEFT JOIN saturn_crm.callcenter AS c ON r.callcenter_id = c.id ORDER BY r.`path`')
        rows = cursor.fetchall()
        path = int(rows[0][0])
        count = 1
        call_dates = [rows[0][1].date()]
        self.report_rez = {}
        for i, row in enumerate(rows):
            monodate = ''
            if i:
                if int(row[0]) == path:
                    # Папка не поменялась
                    count += 1
                    if row[1].date() not in call_dates:
                        call_dates.append(row[1].date())
                else:
                    # Следующая папка
                    if count == 0:
                        monodate = ''
                    elif count == 1:
                        monodate = 'начато'
                    elif len(call_dates) == 1:
                        monodate = call_dates[0].strftime('%d.%m.%y')
                    elif len(call_dates) == 2:
                        if abs(datetime.combine(call_dates[0],time(0,0,0,0)) - \
                                                    datetime.combine(call_dates[1],time(0,0,0,0))) < timedelta(days=2):
                            monodate = datetime.combine(call_dates[0],time(0,0,0,0)).strftime('%d-') + \
                                                   datetime.combine(call_dates[1],time(0,0,0,0)).strftime('%d.%m.%y')
                        else:
                            monodate = 'МУЛЬТИ'
                    else:
                        monodate = 'МУЛЬТИ'
                    self.report_rez[path] = monodate
                    call_dates = [row[1].date()]
                    path = int(row[0])
                    count = 1
        if count == 0:
            monodate = ''
        elif count == 1:
            monodate = 'начато'
        elif len(call_dates) == 1:
            monodate = call_dates[0].strftime('%d.%m.%y')
        elif len(call_dates) == 2:
            if datetime.combine(call_dates[0], time(0, 0, 0, 0)) - \
                    datetime.combine(call_dates[1], time(0, 0, 0, 0)) < timedelta(days=2):
                monodate = datetime.combine(call_dates[0], time(0, 0, 0, 0)).strftime('%d-') + \
                           datetime.combine(call_dates[1], time(0, 0, 0, 0)).strftime('%d.%m.%y')
            else:
                monodate = 'МУЛЬТИ'
        else:
            monodate = 'МУЛЬТИ'
        self.report_rez[path] = monodate
        # Строим нити
        dbconn = MySQLConnection(**self.dbconfig_crm)
        cursor = dbconn.cursor()
        # Распознанные файлы, отсортированные по дате
        cursor.execute('SELECT r.`path`, c.inserted_date FROM lekarh.alone_remont AS r '
                       'LEFT JOIN saturn_crm.callcenter AS c ON r.callcenter_id = c.id '
                       'ORDER BY c.inserted_date, r.`path`')
        rows = cursor.fetchall()
        path = int(rows[0][0])
        self.threads = []
        for i, row in enumerate(rows):
            if i:
                if int(row[0]) == path:
                    # Папка не поменялась
                    if not len(self.threads):
                        # Если вообще ни одной, то создаем первую нить
                        self.threads.append({'start': row[1], 'end': row[1], 'maxPath': int(row[0]),
                                        'pathsDates': {int(row[0]): {row[1]: int(row[0])}}})
                    else:
                        threadUpdated = False
                        for j, thread in enumerate(self.threads):
                            if row[1].date() != thread['end'].date() or int(row[0]) != thread['maxPath']:
                                # Дата и последняя папка не совпадает?
                                if row[1] > thread['end'] and (thread['end'] + timedelta(days=6)) > row[1] and \
                                        int(row[0]) >= thread['maxPath']:
                                    # Меньше 15 дней и директория та же или увеличилась? Добавляем в рамках этой нити
                                    self.threads[j]['end'] = row[1]
                                    self.threads[j]['maxPath'] = int(row[0])
                                    if self.threads[j]['pathsDates'].get(int(row[0]), None):
                                        self.threads[j]['pathsDates'][int(row[0])][row[1]] = int(row[0])
                                    else:
                                        self.threads[j]['pathsDates'][int(row[0])] = {row[1]: int(row[0])}
                                    threadUpdated = True
                            else:
                                threadUpdated = True
                                break
                        if not threadUpdated:
                            # Создаем новую нить
                            self.threads.append({'start': row[1], 'end': row[1], 'maxPath': int(row[0]),
                                            'pathsDates': {int(row[0]): {row[1]: int(row[0])}}})
                else:
                    # Следующая папка
                    if not len(self.threads):
                        # Если вообще ни одной, то создаем первую нить
                        self.threads.append({'start': row[1], 'end': row[1], 'maxPath': int(row[0]),
                                        'pathsDates': {int(row[0]): {row[1]: int(row[0])}}})
                    else:
                        threadUpdated = False
                        for j, thread in enumerate(self.threads):
                            if row[1].date() != thread['end'].date():
                                # Дата и последняя папка не совпадает?
                                if row[1] > thread['end'] and (thread['end'] + timedelta(days=6)) > row[1] and \
                                        int(row[0]) >= thread['maxPath']:
                                    # Меньше 15 дней и директория та же или увеличилась? Добавляем в рамках этой нити
                                    self.threads[j]['end'] = row[1]
                                    self.threads[j]['maxPath'] = int(row[0])
                                    if self.threads[j]['pathsDates'].get(int(row[0]), None):
                                        self.threads[j]['pathsDates'][int(row[0])][row[1]] = int(row[0])
                                    else:
                                        self.threads[j]['pathsDates'][int(row[0])] = {row[1]: int(row[0])}
                                    threadUpdated = True
                            else:
                                threadUpdated = True
                                break
                        if not threadUpdated:
                            # Создаем новую нить
                            # Создаем новую нить
                            self.threads.append({'start': row[1], 'end': row[1], 'maxPath': int(row[0]),
                                            'pathsDates': {int(row[0]): {row[1]: int(row[0])}}})

        keys = []
        for i in range(0, 10):
            keys.append(str(i))
        hkeys = []
        for i in range(0, 546):
            hkeys.append(str(i))
        self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
        self.twRez.setRowCount(546)  # Кол-во строк из таблицы
        for j in range(0, 546):
            for k in range(0, 10):
                if self.report_rez.get(j * 10 + k, None):
                    self.twRez.setItem(j, k, QTableWidgetItem(self.report_rez[j * 11 + k]))
                else:
                    self.twRez.setItem(j, k, QTableWidgetItem('нетинф'))
        # Устанавливаем заголовки таблицы
        self.twRez.setHorizontalHeaderLabels(keys)
        # Устанавливаем заголовки таблицы
        self.twRez.setVerticalHeaderLabels(hkeys)
        # Устанавливаем выравнивание на заголовки
        self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twRez.resizeColumnsToContents()
        self.clbReport2xlsx.setEnabled(True)


    def click_clbRefreshReportOld(self):
        dbconn = MySQLConnection(**self.dbconfig_alone)
        cursor = dbconn.cursor()
        cursor.execute('SELECT client_id, path, file FROM alone_connect')
        rows = cursor.fetchall()
        #rows = [('39f07f6d-16e7-11e8-86b5-5254004b76e6', '1', 'f658161664'), ('8113962c-16b8-11e8-86b5-5254004b76e6', '2', 'f659636224'), ('42610b96-16ea-11e8-86b5-5254004b76e6', '2', 'f659652608'), ('f51b5fd6-178d-11e8-86b5-5254004b76e6', '2', 'f687063040'), ('8c95a423-bbc5-11e6-b8cb-20cf300dec24', '6', 'f1083621376'), ('820c39ab-178f-11e8-86b5-5254004b76e6', '6', 'f1083621376'), ('3c0652e8-1809-11e8-81ec-5254004b76e6', '7', 'f2880913408'), ('d2d14811-18a0-11e8-81ec-5254004b76e6', '8', 'f3592290304'), ('b77d6b73-04ae-11e7-9f62-5254004b76e6', '9', 'f3592732672'), ('d30b3605-180c-11e8-81ec-5254004b76e6', '9', 'f3592732672'), ('525e4b86-d737-11e6-aa92-20cf300dec24', '9', 'f3712696320'), ('d8fa8330-178f-11e8-86b5-5254004b76e6', '9', 'f3712696320'), ('d3950684-fd16-11e8-8408-000c290cfc84', '25', 'f3909402624'), ('2c80db53-fddc-11e8-8408-000c290cfc84', '25', 'f3899719680'), ('c3b71aed-1da7-11e7-8786-5254004b76e6', '150', 'f2267316224'), ('f2c3a29d-156f-11e8-9039-5254004b76e6', '150', 'f2267807744'), ('d758b226-1808-11e8-81ec-5254004b76e6', '150', 'f2267807744'), ('3830045c-047b-11e9-a9ee-000c290cfc84', '150', 'f2267807744'), ('e4eba1e7-9010-11e7-8989-5254004b76e6', '3', 'f920420352'), ('f8262d02-ec50-11e7-897e-5254004b76e6', '3', 'f920420352')]
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
        #rows = [('ЮРЬЕВ', 'МИХАИЛ', 'АНАТОЛЬЕВИЧ', 'Омская обл, г Омск, Советский округ, ул Химиков, д 16, кв 61', 'Омская обл, г Омск, Советский округ, ул Химиков, д 16, кв 61', 79831198154, '', datetime(2018, 12, 12, 12, 54, 2), '2c80db53-fddc-11e8-8408-000c290cfc84', date(1989, 6, 15))]
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
        self.report_rez = {}
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
                    self.report_rez[path] = datetime.combine(data,time(0,0,0,0)).strftime('%d.%m.%y')
                elif len(report2phones[path]) > 1:
                    self.report_rez[path] = 'МУЛЬТИ'
                else:
                    self.report_rez[path] = 'начато'
                break
        keys = []
        for i in range(0, 10):
            keys.append(str(i))
        hkeys = []
        for i in range(0, 546):
            hkeys.append(str(i))
        self.twRez.setColumnCount(len(keys))  # Устанавливаем кол-во колонок
        self.twRez.setRowCount(546)  # Кол-во строк из таблицы
        for j in range(0, 546):
            for k in range(0, 10):
                if self.report_rez.get(j * 10 + k, None):
                    self.twRez.setItem(j, k, QTableWidgetItem(self.report_rez[j * 10 + k]))
                else:
                    self.twRez.setItem(j, k, QTableWidgetItem('нетинф'))
        # Устанавливаем заголовки таблицы
        self.twRez.setHorizontalHeaderLabels(keys)
        # Устанавливаем заголовки таблицы
        self.twRez.setVerticalHeaderLabels(hkeys)
        # Устанавливаем выравнивание на заголовки
        self.twRez.horizontalHeaderItem(0).setTextAlignment(Qt.AlignCenter)
        # делаем ресайз колонок по содержимому
        self.twRez.resizeColumnsToContents()
        self.clbReport2xlsx.setEnabled(True)

    def click_clbReport2xlsx(self):
        wb_log = openpyxl.Workbook(write_only=True)
        ws_log = wb_log.create_sheet('Отчет')
        keys = []
        for i in range(-1, 10):
            keys.append(str(i))
        ws_log.append(keys)
        for i in range(0, 546):
            xlsx_str = []
            xlsx_str.append(str(i))
            for j in range(0, 10):
                if self.report_rez.get(i * 10 + j, None):
                    xlsx_str.append(self.report_rez[i * 10 + j])
                else:
                    xlsx_str.append('нетинф')
            ws_log.append(xlsx_str)
        wb_log.save('Отчет.xlsx')
        return

