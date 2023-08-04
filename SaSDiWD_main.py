#! python3
# UA: Програма ітерує файли у форматі  .docx з вибраної користувачем директорії та створює табілицю  з даними о цих файлах з якої користувач може відкрити цей файл
# EN: The program iterates files in .docx format from the directory selected by the user and creates a table with data about these files from which the user can open this file

import os
import re
import subprocess
import sys
import sqlite3
import os
import docx
from PyQt5 import QtCore, QtGui, QtWidgets
from Main_Window import *
from datetime import datetime

os.system('chcp 65001')  # кодіровка термінала віндовс для відкриття

# Сегмент коду який відповідає за створення БД

# створюємо БД
con = sqlite3.connect("filters.db")  
# створюємо курсор
cur = con.cursor()  # створив курсор
cur.execute("CREATE TABLE IF NOT EXISTS key (name,key1,key2,key3)")
cur.execute("CREATE TABLE IF NOT EXISTS way (npp, dir)")
con.close


class MyWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton_2.clicked.connect(self.open_folder)
        self.ui.pushButton_4.clicked.connect(self.open_file)
        self.ui.pushButton_5.clicked.connect(self.create_new_filter)
        self.ui.pushButton_3.clicked.connect(
            self.create_fin_list)  # Кнопка пошуку
        self.ui.pushButton.clicked.connect(self.delete_filter)

        self.directory = ''
        self.filepath = ''
        self.list_of_fname = []

        name_fitr_in_BD = cur.execute("SELECT name FROM key ")
        name_BD = name_fitr_in_BD.fetchall()
        for name_filter in name_BD:
            for name in name_filter:
                self.ui.listWidget.addItem(name)
                self.list_of_fname.append(name)
                self.ui.comboBox.addItem(name)

        con.close


    def create_fin_list(self):
        con = sqlite3.connect("filters.db")
        cur = con.cursor()
        cur.execute("DELETE  FROM way")
        con.commit()
        self.ui.listWidget_2.clear()
        first_line = 'НПП' + '  | ' + 'Дата останьої модифікації' + \
            ' ' * 8 + '|   ' + "Назва файлу"
        self.ui.listWidget_2.addItem(first_line)
        Npp = 1
        try:
            for file_name in os.listdir(os.path.join(self.directory)):

                if file_name.endswith('.docx'):
                    appropriate_file = f'{self.directory}/{file_name}'
                    appropriate_file_text = self.get_text_from_word_file(
                        appropriate_file)
                    Result = self.find_keyword_in_files(appropriate_file_text)

                    if Result:
                        list_number = self.find_number_of_name_file(
                            appropriate_file)
                        mod_data = self.last_modification_date(
                            appropriate_file)
                        line = str(Npp).ljust(3) + ' '*(7-len(str(Npp))) + \
                            '| ' + mod_data + ' '*20 + '|   ' + file_name
                        # line = f'{Npp.rjust}          {file_name}          {mod_data}          {list_number}'
                        self.ui.listWidget_2.addItem(line)
                        tuple_keys = (str(Npp), appropriate_file)
                        cur.execute(
                            "INSERT INTO way(npp, dir) VALUES (?,?)", tuple_keys)
                        con.commit()
                        Npp += 1

        except FileNotFoundError:
            QtWidgets.QMessageBox.warning(
                self, "Увага", "Тека для пошуку не вибрана")
        except:
            None
        cur.close

    def get_text_from_word_file(self, path):  # Получення повного текту з файла
        document = docx.Document(path)
        fullText = []
        for para in document.paragraphs:
            fullText.append(para.text)
        Text = '/n'.join(fullText)
        return Text

    def find_number_of_name_file(self, filename):   # Пошу номерів лістів
        names_list = []
        letter_index = re.compile(
            r'(\d\d-\d\d\d\d-\d\d-\d\d)|(\d\d-\d\d\d\d-\d\d)|(\d\d-\d\d-\d\d)|(\d\d-\d\d-\d\d)|(\d\d-\d\d) ')
        seash_var = letter_index.search(filename)
        if seash_var == None:
            names_list.append('Відсутній номер договору')
        else:
            seash_var = seash_var.group()
            names_list.append(seash_var)
        return str(names_list)

    def find_keyword_in_files(self, Text):  # Пошук по ключевим словам
        filter_name = self.ui.comboBox.currentText()
        con = sqlite3.connect("filters.db")  # створили БД
        cur = con.cursor()  # створив курсор
        kwp_all_BD = cur.execute(
            f"SELECT key1, key2, key3 FROM key WHERE name = '{filter_name}'")
        kwp_all = kwp_all_BD.fetchall()
        kwp1 = kwp_all[0][0]
        kwp2 = kwp_all[0][1]
        kwp3 = kwp_all[0][2]
        con.close
        if re.search(kwp1, Text) and re.search(kwp2, Text) and re.search(kwp3, Text):
            return True
        else:
            return None

    def create_new_filter(self):

        name_filter = self.ui.lineEdit.text()
        kerword1 = self.ui.lineEdit_2.text()
        ketword2 = self.ui.lineEdit_3.text()
        keyword3 = self.ui.lineEdit_4.text()
        tuple_keys = (name_filter, kerword1, ketword2, keyword3)

        if not name_filter:
            QtWidgets.QMessageBox.warning(
                self, "Увага", "Строка з ім`ям фільтра порожня")
        else:
            if name_filter in self.list_of_fname:
                QtWidgets.QMessageBox.warning(
                    self, "Увага", "Фільтр з таким ім'ям вже існує")
            else:

                # повторне підключення до БД
                con = sqlite3.connect("filters.db")
                cur = con.cursor()
                cur.execute("INSERT INTO key VALUES (?,?,?,?)", tuple_keys)
                con.commit()

                name_fitr_in_BD = cur.execute("SELECT name FROM key ")
                name_BD = name_fitr_in_BD.fetchall()
                self.ui.listWidget.clear()
                for name_filter in name_BD:
                    for name in name_filter:
                        self.ui.listWidget.addItem(name)
                        self.list_of_fname.append(name)
                self.ui.lineEdit.clear()
                self.ui.lineEdit_2.clear()
                self.ui.lineEdit_3.clear()
                self.ui.lineEdit_4.clear()
                self.ui.comboBox.addItem(name)
                con.close()

    def open_folder(self):  # відкритя потрібної теки з файлами
        self.directory = QtWidgets.QFileDialog.getExistingDirectory(self)
        self.ui.label_2.setText(f'Пошук у теці: {self.directory}')

    

    def open_file(self):
        info = self.ui.listWidget_2.currentItem().text()
        info_list = str(info).split()
        number = tuple(info_list[:1])
        con = sqlite3.connect("filters.db")  # створили БД
        cur = con.cursor()  # створив курсор
        bd_start = cur.execute("SELECT dir FROM way WHERE npp = ?  ", number)
        start = bd_start.fetchone()
        start = start[0]
        start = os.path.realpath(start)
        os.startfile(start)
        # subprocess.Popen(['start', start], shell=True)
        con.close

    def delete_filter(self):

        try:
            selected = self.ui.listWidget.currentItem().text()
            list_selected = []
            list_selected.append(str(selected))
            tuple_selected = tuple(list_selected)
            print(tuple_selected)
            con = sqlite3.connect("filters.db")  # повторне підключення до БД
            cur = con.cursor()
            cur.execute("DELETE FROM key WHERE name = ?",  tuple_selected)
            con.commit()
            name_fitr_in_BD = cur.execute("SELECT name FROM key ")
            name_BD = name_fitr_in_BD.fetchall()
            self.ui.listWidget.clear()
            self.ui.comboBox.clear()
            self.list_of_fname.remove(selected)
            for name_filter in name_BD:
                for name in name_filter:
                    self.ui.listWidget.addItem(name)
                    self.ui.comboBox.addItem(name)
            con.close()
        except AttributeError:
            return None

    def last_modification_date(self, path):
        mod_data = str(datetime.fromtimestamp(os.path.getmtime(path)))
        mod_data = mod_data[:19]
        return mod_data


# точка входу
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())
