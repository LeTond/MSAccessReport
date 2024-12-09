import pyodbc


class AccessBack:
    def __init__(self, filepath, first_data, last_data):
        self.counter = 0
        self.filepath = filepath
        self.first_data = first_data
        self.last_data = last_data

    def pyodbc_cursor(self):
        conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                              f'DBQ={self.filepath};')
        cursor = conn.cursor()
        return cursor

    @staticmethod
    def exist_categories() -> list:
        list_ = ['Пл', 'ВМП', 'КА', 'Дог', 'Сотр', 'Суб', 'ДМС', 'Грант', 'Наука']
        return list_

    def all_studies(self) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#').fetchall()
        return len(list_data)

    def all_studies_with_contrast(self) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Контраст=True '
                    f'AND Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#').fetchall()
        return len(list_data)

    def studies(self, category) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Категория=? '
                    f'AND Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#', category).fetchall()
        return len(list_data)

    def studies_with_contrast(self, category) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Контраст=? '
                    f'AND Категория=? '
                    f'AND Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#', (True, category)).fetchall()
        return len(list_data)

    def amb_oms_counter_with_contrast(self) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Контраст=? '
                    f'AND Категория=? '
                    f'AND Отделение=? '
                    f'AND Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#', (True, 'ОМС', 'Амб')).fetchall()
        return len(list_data)

    def amb_oms_studies(self) -> int:
        list_data = self.pyodbc_cursor(). \
            execute(f'SELECT * FROM Журнал WHERE '
                    f'Категория=? '
                    f'AND Отделение=? '
                    f'AND Дата_исследования>#{self.first_data}# '
                    f'AND Дата_исследования<#{self.last_data}#', ('ОМС', 'Амб')).fetchall()
        return len(list_data)

    def week_report(self) -> dict:
        diction_week = {
            'Всего': self.all_studies() + self.all_studies_with_contrast()
        }
        for cat in self.exist_categories():
            diction_week[f'{cat}'] = self.studies_with_contrast(cat) + \
                                     self.studies(cat)

        for cash in ['Пл', 'Сотр', 'ДМС', 'Дог']:
            self.counter += self.studies(cash) + \
                            self.studies_with_contrast(cash)

        diction_week['Всего платно'] = self.counter

        diction_week['ОМС стац'] = self.studies_with_contrast('ОМС') + \
                                   self.studies('ОМС') - \
                                   self.amb_oms_counter_with_contrast() - \
                                   self.amb_oms_studies()

        diction_week['Всего стац'] = diction_week['ВМП'] + diction_week['ОМС стац']

        diction_week['С контрастом'] = self.all_studies_with_contrast()

        diction_week['ОМС амб'] = self.amb_oms_counter_with_contrast() + \
                                  self.amb_oms_studies()
        return diction_week
#
#
# if __name__ == '__main__':
#     acc = AccessBack('C:/Users/User/Desktop/Statistika/Base_2020.accdb')
#     print(acc.all('2020-01-20', '2020-09-25'))
