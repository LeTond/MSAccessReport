from datetime import date
from access import AccessBack
from report import Report

from PyQt5.QtWidgets import QApplication, QFileDialog
from PyQt5 import QtWidgets, uic
import sys


class Main(QtWidgets.QMainWindow, QtWidgets.QDateEdit):

    def __init__(self, parent=None):
        QtWidgets.QMainWindow.__init__(self, parent)
        uic.loadUi("gui/StatisticGUI.ui", self)

        self.ChooseAccess.clicked.connect(self.get_directory)
        self.btn_Week.clicked.connect(self.week_report)
        self.btn_Month.clicked.connect(self.month_report)
        self.btn_Quarter.clicked.connect(self.quarter_report)
        self.btn_Year.clicked.connect(self.year_report)
        self.btn_Close_Main.clicked.connect(self.closeEvent)
        self.filepath = self.get_directory()

    def get_directory(self) -> str:
        """
        Pick requiring microsoft access base: file_name.accdb in your directory
        :return: filepath_to_directory_with_file/filename.accdb
        """
        filepath, filetype = QFileDialog.getOpenFileName(self,
                                                         "Выбрать файл",
                                                         ".",
                                                         "Text Files(*.accdb);;\
                                                         ;;All Files(*)")
        self.set_current_date()
        return filepath

    def set_current_date(self):
        self.dateFrom.setDate(date.today())
        self.dateTo.setDate(date.today())

    def date_from(self):
        date_from = self.dateFrom.dateTime()
        return date_from

    def date_to(self):
        date_to = self.dateTo.dateTime().addDays(1)
        return date_to

    def real_date_to(self):
        real_date_to = self.dateTo.dateTime()
        return real_date_to

    def week_report(self):
        """
        Run methods to create week report
        :return: None
        """
        access = AccessBack(self.filepath, f"{self.date_from().toString('MM/dd/yyyy')}",
                            f"{self.date_to().toString('MM/dd/yyyy')}")
        rp = Report()
        date_period = f"{self.date_from().toString('dd.MM.yyyy')} - {self.real_date_to().toString('dd.MM.yyyy')}"
        dict_ = access.week_report()
        rp.generate_week_document(dict_, date_period)

    def month_report(self):
        """
        Run methods to create month report
        :return: None
        """
        access = AccessBack(self.filepath, f"{self.date_from().toString('MM/dd/yyyy')}",
                            f"{self.date_to().toString('MM/dd/yyyy')}")
        rp = Report()
        date_period = f"{self.date_from().toString('dd.MM.yyyy')} - {self.real_date_to().toString('dd.MM.yyyy')}"
        dict_ = access.month_report()
        rp.generate_month_document(dict_, date_period)

    def quarter_report(self):
        """
        Run methods to create month report
        :return: None
        """
        access = AccessBack(self.filepath, f"{self.date_from().toString('MM/dd/yyyy')}",
                            f"{self.date_to().toString('MM/dd/yyyy')}")
        rp = Report()
        date_period = f"{self.date_from().toString('dd.MM.yyyy')} - {self.real_date_to().toString('dd.MM.yyyy')}"
        dict_ = access.quarter_report()
        rp.generate_quarter_document(dict_, date_period)

    def year_report(self):
        """
        Run methods to create month report
        :return: None
        """
        access = AccessBack(self.filepath, f"{self.date_from().toString('MM/dd/yyyy')}",
                            f"{self.date_to().toString('MM/dd/yyyy')}")
        rp = Report()
        date_period = f"{self.date_from().toString('dd.MM.yyyy')} - {self.real_date_to().toString('dd.MM.yyyy')}"
        dict_ = access.year_report()
        rp.generate_year_document(dict_, date_period)

    def closeEvent(self, event):
        result = QtWidgets.QMessageBox.question(self,
                                                "Подтверждение закрытия окна",
                                                "Вы действительно хотите закрыть окно?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                QtWidgets.QMessageBox.No)
        if result == QtWidgets.QMessageBox.Yes:
            app.quit()
        else:
            pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mn = Main()
    mn.show()

    sys.exit(app.exec())