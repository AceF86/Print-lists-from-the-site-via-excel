import sqlite3
import subprocess
import sys
import os

from functools import partial

from PyQt5.QtCore import QSettings, QTimer, QSize, QLocale
import PyQt5.QtWidgets as qtw
from PyQt5 import QtCore
import PyQt5.QtGui as qtg
import win32api
import win32print
import exel_maker
import jsonData


def resource_path(relative_path):
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


path = resource_path("data/list.db")
conn = sqlite3.connect(path)
c = conn.cursor()
c.execute("CREATE TABLE IF NOT EXISTS list_judge(judge1)")
conn.commit()


class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()

        self.saveSetting()
        try:
            self.move(self.settings.value("window position"))
        except:
            pass
        self.setWindowTitle("Список справ")
        self.setWindowFlag(QtCore.Qt.WindowMinimizeButtonHint, False)
        self.l_place_value = self.settings.value("name")
        self.l_place = qtw.QLineEdit(self)
        self.l_place.setText(self.l_place_value)
        self.l_dateSel = qtw.QLabel()

        self.comboBox = qtw.QComboBox(self)
        path = resource_path("data/list.db")
        self.conn = sqlite3.connect(path)
        self.c = self.conn.cursor()

        try:
            self.result = self.c.execute("SELECT * FROM list_judge")
            self.rows = self.result.fetchall()
            for row in self.rows:
                self.comboBox.addItem(row[0])
            self.conn.commit()
        except:
            pass

        self.c.close()

        form_layout = qtw.QFormLayout()
        self.setLayout(form_layout)
        self.layout1 = qtw.QLabel()
        form_layout.addRow(self.layout1)
        form_layout.addRow(qtw.QLabel())
        form_layout.addRow("", self.comboBox)
        form_layout.addRow(qtw.QLabel())
        form_layout.addRow(
            "                                     "
            "                                           "
            "       ",
            self.l_place,
        )
        form_layout.addRow(qtw.QLabel())
        form_layout.addRow(qtw.QLabel())

        btn_excel = qtw.QPushButton("  Відкрити", self)
        btn_excel.clicked.connect(lambda: self.open_exel())
        btn_excel.move(277, 140)
        btn_excel.resize(112, 27)
        path = resource_path("data/_xls_icon2.png")
        btn_excel.setIcon(qtg.QIcon(path))
        btn_excel.setIconSize(QSize(20, 20))

        btn_print = qtw.QPushButton("  Друк", self)
        btn_print.clicked.connect(lambda: self.print_exel())
        btn_print.move(277, 175)
        btn_print.resize(112, 27)
        path = resource_path("data/printer icon3.png")
        btn_print.setIcon(qtg.QIcon(path))
        btn_print.setIconSize(QSize(22, 22))

        btn_menu = qtw.QPushButton("", self)
        btn_menu.clicked.connect(lambda: self.show_AnotherWindow())
        btn_menu.move(277, 9)
        btn_menu.resize(35, 35)
        path = resource_path("data/_tools.png")
        btn_menu.setIcon(qtg.QIcon(path))
        btn_menu.setIconSize(QSize(30, 30))

        btn_Info = qtw.QPushButton("", self)
        btn_Info.clicked.connect(lambda: self.messageBox())
        btn_Info.move(374, 9)
        btn_Info.resize(15, 15)
        path = resource_path("data/info_black.png")
        btn_Info.setIcon(qtg.QIcon(path))
        btn_Info.setIconSize(QSize(15, 20))

        self.btn = qtw.QPushButton(self)
        self.btn.clicked.connect(partial(self.dowlou_json))
        self.btn.resize(0, 0)

        self.l_date = qtw.QCalendarWidget(self)
        self.l_date.setGridVisible(True)
        self.l_date.clicked[QtCore.QDate].connect(self.showDate)
        self.l_date.setVerticalHeaderFormat(qtw.QCalendarWidget.NoVerticalHeader)
        self.l_date.setLocale(QLocale(QLocale.Ukrainian))
        self.l_date.setFont(qtg.QFont("Sanserif", 8))
        self.l_date.setFixedSize(QSize(260, 192))
        self.l_date.move(10, 10)

        self.lbl = qtw.QLabel(self)
        date = self.l_date.selectedDate()
        self.lbl.setText(date.toString("dd.MM.yyyy"))
        self.lbl.resize(0, 0)

        jud = qtw.QLabel(self)
        jud.setText("Суддя")
        jud.setStyleSheet("QLabel {font-size: 13pt;}")
        jud.move(317, 22)

        secretary = qtw.QLabel(self)
        secretary.setText("Секретар")
        secretary.setStyleSheet("QLabel {font-size: 10pt;}")
        secretary.move(278, 68)

        self.comboBox_state = self.settings.value("comboBox_judge", 4)
        self.comboBox.setCurrentIndex(self.comboBox_state)
        self.comboBox.activated.connect(lambda: self.saveSetting())

    def showDate(self, date):
        self.lbl.setText(date.toString("dd.MM.yyyy"))

    def print_exel(self):
        try:
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data_pr.json",
            )

            self.timerMessageBox()
            win32api.ShellExecute(
                0,
                "print",
                "book.xlsx",
                '/d: "%s"' % win32print.GetDefaultPrinter(),
                ".",
                7,
            )

        except IOError as e:
            DETACHED_PROCESS = 0x00000008
            subprocess.call('wmic process where name="EXCEL.EXE" delete', creationflags=DETACHED_PROCESS)
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data_pr.json",
            )
            self.timerMessageBox()
            win32api.ShellExecute(
                0,
                "print",
                "book.xlsx",
                '/d: "%s"' % win32print.GetDefaultPrinter(),
                ".",
                7,
            )

        except Exception as ex:
            DETACHED_PROCESS = 0x00000008
            subprocess.call('wmic process where name="EXCEL.EXE" delete', creationflags=DETACHED_PROCESS)
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data/data_pr.json",
            )
            self.timerMessageBox()
            win32api.ShellExecute(
                0,
                "print",
                "book.xlsx",
                '/d: "%s"' % win32print.GetDefaultPrinter(),
                ".",
                7,
            )

    def open_exel(self):
        try:
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data_pr.json",
            )
            win32api.ShellExecute(0, "open", "book.xlsx", None, ".", 0)

        except IOError as e:
            DETACHED_PROCESS = 0x00000008
            subprocess.call('wmic process where name="EXCEL.EXE" delete', creationflags=DETACHED_PROCESS)
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data_pr.json",
            )
            win32api.ShellExecute(0, "open", "book.xlsx", None, ".", 0)

        except Exception as ex:
            DETACHED_PROCESS = 0x00000008
            subprocess.call('wmic process where name="EXCEL.EXE" delete', creationflags=DETACHED_PROCESS)
            exel_maker.create_exel(
                self.lbl.text(),
                self.comboBox.currentText(),
                self.l_place.text(),
                "data/data_pr.json",
            )
            win32api.ShellExecute(0, "open", "book.xlsx", None, ".", 0)

    def dowlou_json(self):

        try:
            jsonData.makeJsonData()
        except:
            print("Error")

    def messageBox(self):
        qtw.QMessageBox.about(
            self,
            "Інформація",
            "Список справ без інтернету не працює правильно.\n"
            "Но вразі відсутності інтернету буде запас на\n3 дні інформації.\n"
            "Працює на  64 bit версії Windows.\n"
            "Обов'язково потрібно установити Microsoft Excel.\n\n"
            "Програму було створено Мордованець Русланом\n\n"
            "E-mail: gifler@me.com\n"
            "2022 рік.",
        )

    # ============================= check_box ====================================== #

    def saveSetting(self):
        self.settings = QSettings("Main menu", "window location")
        self.settings_windows1 = QSettings("main window", "web_win location")

    def closeEvent(self, event):
        self.settings.setValue("window position", self.pos())
        self.settings.setValue("name", self.l_place.text())
        self.settings.setValue("comboBox_judge", self.comboBox.currentIndex())

    # =================================== open Window ========================================= #

    def show_AnotherWindow(self):
        self.dialog = AnotherWindow()
        path = resource_path("data/folder.ico")
        self.dialog.setWindowIcon(qtg.QIcon(path))
        self.dialog.show()
        self.close()

    def timerMessageBox(self):
        messagebox = TimerMessageBox(3, self)
        messagebox.setWindowTitle("Повідомлення")
        messagebox.exec_()


# =========================== Another Window ============================ #


class AnotherWindow(qtw.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Внести данні")
        self.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint, False)
        self.setFixedSize(250, 124)

        self.l_recipient_1 = qtw.QLineEdit(self)

        self.comboBox2 = qtw.QComboBox(self)
        path = resource_path("data/list.db")
        self.conn = sqlite3.connect(path)
        self.c = self.conn.cursor()

        try:
            self.result = self.c.execute("SELECT * FROM list_judge")
            self.rows = self.result.fetchall()
            for row in self.rows:
                self.comboBox2.addItem(row[0])
            self.conn.commit()
        except:
            pass

        form_layout2 = qtw.QFormLayout()
        self.setLayout(form_layout2)

        form_layout2.addRow("Внести суддю       ", self.l_recipient_1)
        form_layout2.addRow(qtw.QLabel())
        form_layout2.addRow(qtw.QLabel(), self.comboBox2)
        form_layout2.addRow(qtw.QLabel())

        push_btn = qtw.QPushButton("  Додати", self)
        push_btn.clicked.connect(lambda: self.save_items())
        push_btn.move(9, 90)
        push_btn.resize(100, 27)
        path = resource_path("data/user-add.png")
        push_btn.setIcon(qtg.QIcon(path))
        push_btn.setIconSize(QSize(20, 20))

        btn_delete = qtw.QPushButton("  Видалити", self)
        btn_delete.clicked.connect(lambda: self.delete_items())
        btn_delete.move(140, 90)
        btn_delete.resize(100, 27)
        path = resource_path("data/user-delete.png")
        btn_delete.setIcon(qtg.QIcon(path))
        btn_delete.setIconSize(QSize(20, 20))

        nameprice = qtw.QLabel(self)
        nameprice.setText("Видалити суддю")
        nameprice.move(10, 50)

    def save_items(self):
        try:
            path = resource_path("data/list.db")
            conn = sqlite3.connect(path)
            c = conn.cursor()
            c.execute(
                f"INSERT INTO list_judge(judge1) VALUES('{self.l_recipient_1.text()}');"
            )
            conn.commit()
            qtw.QMessageBox.information(
                self, "Збережено", f"{self.l_recipient_1.text()}", qtw.QMessageBox.Ok
            )
            self.l_recipient_1.clear()
            self.updat_box()
        except Exception as ex:
            qtw.QMessageBox.critical(self, "Помилка", f"{ex}", qtw.QMessageBox.Ok)
            self.l_recipient_1.clear()

    def delete_items(self):
        try:
            path = resource_path("data/list.db")
            conn = sqlite3.connect(path)
            c = conn.cursor()
            c.execute(
                f'DELETE FROM list_judge WHERE judge1="{self.comboBox2.currentText()}"'
            )
            conn.commit()
            qtw.QMessageBox.information(
                self, "Видалено", f"{self.comboBox2.currentText()}", qtw.QMessageBox.Ok
            )
            self.updat_box()
        except Exception as ex:
            qtw.QMessageBox.critical(self, "Помилка", f"{ex}", qtw.QMessageBox.Ok)

    def updat_box(self):
        self.comboBox2.clear()
        self.result = self.c.execute("SELECT * FROM list_judge")
        self.rows = self.result.fetchall()
        for row in self.rows:
            self.comboBox2.addItem(row[0])
        self.conn.commit()

    def updat_box2(self):
        self.close()
        self.dialog = MainWindow()
        self.dialog.setWindowTitle("Список справ")
        path = resource_path("data/folder.ico")
        self.dialog.setWindowIcon(qtg.QIcon(path))
        self.dialog.setFixedSize(400, 210)
        self.dialog.show()

    def closeEvent(self, event):
        self.updat_box2()


# ============================ Timer Message Box ============================================ #


class TimerMessageBox(qtw.QMessageBox):
    def __init__(self, timeout=1, parent=None):
        super(TimerMessageBox, self).__init__(parent)
        self.setWindowTitle("Повідомлення")
        path = resource_path("data/folder.ico")
        self.setWindowIcon(qtg.QIcon(path))
        self.time_to_wait = timeout
        self.setText("Друк .....".format(timeout))
        self.setStandardButtons(qtw.QMessageBox.Ok)
        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.changeContent)
        self.timer.start()

    def changeContent(self):
        self.setText("Друк .......".format(self.time_to_wait))
        self.time_to_wait -= 1
        if self.time_to_wait <= 0:
            self.close()

    def closeEvent(self, event):
        self.timer.stop()
        event.accept()


if __name__ == "__main__":
    app = qtw.QApplication(sys.argv)
    mw = MainWindow()
    path = resource_path("data/folder.ico")
    mw.setWindowIcon(qtg.QIcon(path))
    mw.setFixedSize(400, 210)
    mw.show()
    sys.exit(app.exec_())
