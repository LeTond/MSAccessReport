import pandas_access as mdb


class AccessBack:
    def __init__(self, filepath: str, first_data: str, last_data: str):
        self.counter = 0
        self.perinatal_mri = mdb.read_table(filepath, encoding='utf-8', table_name="Журнал")
        self.categories_mri = mdb.read_table(filepath, encoding='utf-8', table_name="Категории")
        self.departments_mri = mdb.read_table(filepath, encoding='utf-8',  table_name="Отделения")

        self.from_data_cond = self.perinatal_mri['Дата исследования'] > first_data  # day/month/year
        self.to_data_cond = self.perinatal_mri['Дата исследования'] < last_data  # day/month/year
        self.contrast_cond = self.perinatal_mri["Контраст"] == '1'
        self.not_contrast_cond = self.perinatal_mri["Контраст"] == '0'
        self.amb_cond = self.perinatal_mri["Отделение"] == "Амб"
        self.not_amb_cond = self.perinatal_mri["Отделение"] != "Амб"
        self.oms_cond = self.perinatal_mri["Категория (ИФ)"] == "ОМС"
        self.dms_cond = self.perinatal_mri["Категория (ИФ)"] == "ДМС"
        self.pay_cond = self.perinatal_mri["Категория (ИФ)"] == "Пл"
        self.free_pay_cond = self.perinatal_mri["Категория (ИФ)"] != "Пл"  # в годовом отчете

        self.all_categories_list = ['Пл', 'ВМП', 'КА', 'Дог', 'Сотр', 'Суб', 'ДМС', 'Грант', 'Наука', 'ОМС']
        self.not_department_list = ['Пл', 'Дог', 'Сотр', 'Суб', 'ДМС', 'Грант', 'Наука', 'Всего']
        self.department_list = ['ОМС', 'ВМП', 'КА']

    def print(self):
        print(self.perinatal_mri)

    def exist_departments(self) -> list:
        """
        Search in DB exist department names
        :return: list with department names
        """
        list_ = []
        for db in self.perinatal_mri['Отделение']:
            if db not in list_:
                list_.append(db)
        return list_

    def exist_studies(self) -> list:
        """
        Search exist studies in DB
        :return: list with study names
        """
        list_ = []
        for db in self.perinatal_mri['Услуга']:
            if db not in list_:
                list_.append(db)
        return list_

    @staticmethod
    def all_studies(filepath) -> list:
        list_ = []
        for db in mdb.read_table(filepath, "Услуги")['Наименование услуги']:
            if db not in list_:
                list_.append(db)
        return list_

    def dep_cond(self, dep):
        dep_cond = self.perinatal_mri['Отделение'] == dep
        return dep_cond

    def cat_cond(self, cat):
        cat_cond = self.perinatal_mri["Категория (ИФ)"] == cat
        return cat_cond

    def week_report(self) -> dict:
        """
        Search in DB studies with filter and create diction with studies and count of it
        :return: diction with statistic information about week studies
        """
        diction = {
            'C контрастом': {},
            'Без контраста': {},
            'Всего': {}
        }

        for cat in self.all_categories_list:
            non_contrast_study = self.perinatal_mri[
                self.from_data_cond & self.to_data_cond & self.cat_cond(cat)].shape[0]
            contrast_study = \
                self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & self.cat_cond(cat) & self.contrast_cond
                    ].shape[0]
            study = non_contrast_study + contrast_study
            diction['C контрастом'].update({f"{cat}": contrast_study})
            diction['Без контраста'].update({f"{cat}": non_contrast_study})
            diction['Всего'].update({f"{cat}": study})

        diction['Без контраста'].update({'Всего': self.perinatal_mri[self.from_data_cond & self.to_data_cond].shape[0]})
        diction['C контрастом'].update(
            {'Всего': self.perinatal_mri[self.from_data_cond & self.to_data_cond & self.contrast_cond].shape[0]})
        diction['Всего'].update({'Всего': diction['C контрастом']['Всего'] + diction['Без контраста']['Всего']})

        diction['Без контраста'].update({'ОМС амб': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.amb_cond & self.oms_cond].shape[0]})
        diction['C контрастом'].update({'ОМС амб': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.amb_cond & self.oms_cond & self.contrast_cond].shape[0]})
        diction['Всего'].update({'ОМС амб': diction['Без контраста']['ОМС амб'] + diction['C контрастом']['ОМС амб']})

        diction['Без контраста'].update({'ОМС стац': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.not_amb_cond & self.oms_cond].shape[0]})
        diction['C контрастом'].update({'ОМС стац': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.not_amb_cond & self.oms_cond & self.contrast_cond].shape[0]})
        diction['Всего'].update(
            {'ОМС стац': diction['Без контраста']['ОМС стац'] + diction['C контрастом']['ОМС стац']})

        return diction

    def month_report(self) -> dict:
        """
        Search in DB studies with filter and create diction with studies and count of it
        :return: diction with statistic information about monthly studies
        """
        diction = {
            'Всего': {},
            'C контрастом': {},
            'В поликлинике': {}
        }

        list_heart = ['Сердце', 'Сосуды Гр. Аорта', 'Сосуды Восх. ОА', 'аорта']
        list_vessels = ['Сосуды Вены ГМ', 'Сосуды ГМ', 'Сосуды шеи', 'Ангио', 'вазоневр']
        list_mediastinum = ['Средостен.']
        list_abdominal = ['БП', 'Надпочечники', 'Железо/Печень', 'Хол-я(МРХПГ)', 'Печень', 'Почки',
                          'Забр', 'железо/сердце']
        list_head = ['ГМ', 'Перфузия', 'Морфометрия', 'Трактография', 'фМРТ', 'Слюнн. Жел.', 'Пазухи', 'Диффузия']
        list_pelvis = ['МТ', 'ГСГ', 'Пельвио', 'Циссы', 'Половые губы', 'рубец на матке']
        list_c_spine = ['ШОП']
        list_th_spine = ['ГОП']
        list_l_spine = ['ПОП']
        list_joints = ['ТбС', 'КПС', 'КС', 'Голен.сустав', 'Мягк. Ткани Плеча', 'Копчик', 'Мягк. Ткани',
                       'Плечевой сустав', 'Мягк. Ткани шеи', 'Мягк. Ткани бедер', 'ВНЧС', 'Мягк. Ткани лица',
                       'Мягк. Ткани копчика', 'Кисть', 'Локоть', 'Мягк. Ткани Гр. Кл.', 'Мягк. Ткани Ягодиц',
                       'Орг рот полости', 'Мягк. Ткани голени', ]
        list_fetus = ['Плод']
        list_placenta = ['Плацента']
        list_other = ['Гипофиз, ХСО', 'Орбиты']

        diction['Всего'].update({'Всего МРТ': self.perinatal_mri[self.from_data_cond & self.to_data_cond].shape[0]})
        diction['C контрастом'].update({'Всего МРТ': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.contrast_cond
            ].shape[0]})
        diction['В поликлинике'].update({'Всего МРТ': self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.oms_cond & self.amb_cond].shape[0]})

        diction['Всего'].update({'ССС': self.service_generator(list_heart)[0]})
        diction['C контрастом'].update({'ССС': self.service_generator(list_heart)[1]})
        diction['В поликлинике'].update({'ССС': self.service_generator(list_heart)[2]})

        diction['Всего'].update({'Легкие и средостения': self.service_generator(list_mediastinum)[0]})
        diction['C контрастом'].update({'Легкие и средостения': self.service_generator(list_mediastinum)[1]})
        diction['В поликлинике'].update({'Легкие и средостения': self.service_generator(list_mediastinum)[2]})

        diction['Всего'].update({'ОБП и ЗП': self.service_generator(list_abdominal)[0]})
        diction['C контрастом'].update({'ОБП и ЗП': self.service_generator(list_abdominal)[1]})
        diction['В поликлинике'].update({'ОБП и ЗП': self.service_generator(list_abdominal)[2]})

        diction['Всего'].update({'Голова': self.service_generator(list_head)[0]})
        diction['C контрастом'].update({'Голова': self.service_generator(list_head)[1]})
        diction['В поликлинике'].update({'Голова': self.service_generator(list_head)[2]})

        diction['Всего'].update({'Малый таз': self.service_generator(list_pelvis)[0]})
        diction['C контрастом'].update({'Малый таз': self.service_generator(list_pelvis)[1]})
        diction['В поликлинике'].update({'Малый таз': self.service_generator(list_pelvis)[2]})

        diction['Всего'].update({'Шейный отдел': self.service_generator(list_c_spine)[0]})
        diction['C контрастом'].update({'Шейный отдел': self.service_generator(list_c_spine)[1]})
        diction['В поликлинике'].update({'Шейный отдел': self.service_generator(list_c_spine)[2]})

        diction['Всего'].update({'Грудной отдел': self.service_generator(list_th_spine)[0]})
        diction['C контрастом'].update({'Грудной отдел': self.service_generator(list_th_spine)[1]})
        diction['В поликлинике'].update({'Грудной отдел': self.service_generator(list_th_spine)[2]})

        diction['Всего'].update({'Пояснично-крестцовый': self.service_generator(list_l_spine)[0]})
        diction['C контрастом'].update({'Пояснично-крестцовый': self.service_generator(list_l_spine)[1]})
        diction['В поликлинике'].update({'Пояснично-крестцовый': self.service_generator(list_l_spine)[2]})

        diction['Всего'].update({'Суставы': self.service_generator(list_joints)[0]})
        diction['C контрастом'].update({'Суставы': self.service_generator(list_joints)[1]})
        diction['В поликлинике'].update({'Суставы': self.service_generator(list_joints)[2]})

        diction['Всего'].update({'Сосуды': self.service_generator(list_vessels)[0]})
        diction['C контрастом'].update({'Сосуды': self.service_generator(list_vessels)[1]})
        diction['В поликлинике'].update({'Сосуды': self.service_generator(list_vessels)[2]})

        diction['Всего'].update({'Плацента': self.service_generator(list_placenta)[0]})
        diction['C контрастом'].update({'Плацента': self.service_generator(list_placenta)[1]})
        diction['В поликлинике'].update({'Плацента': self.service_generator(list_placenta)[2]})

        diction['Всего'].update({'Плода': self.service_generator(list_fetus)[0]})
        diction['C контрастом'].update({'Плода': self.service_generator(list_fetus)[1]})
        diction['В поликлинике'].update({'Плода': self.service_generator(list_fetus)[2]})

        diction['Всего'].update({'Прочие': self.service_generator(list_other)[0]})
        diction['C контрастом'].update({'Прочие': self.service_generator(list_other)[1]})
        diction['В поликлинике'].update({'Прочие': self.service_generator(list_other)[2]})

        return diction

    def service_generator(self, service_list) -> list:
        counter = 0
        counter_with_contrast = 0
        counter_amb_oms = 0
        for service in service_list:
            serv = self.perinatal_mri["Услуга"] == service
            counter += self.perinatal_mri[self.from_data_cond & self.to_data_cond & serv].shape[0]
            counter_with_contrast += \
                self.perinatal_mri[self.from_data_cond & self.to_data_cond & serv & self.contrast_cond].shape[0]
            counter_amb_oms += \
                self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & serv & self.amb_cond & self.oms_cond].shape[0]
        return [counter, counter_with_contrast, counter_amb_oms]

    def count_quarter_studies(self, department, contrast) -> dict:
        dms = self.perinatal_mri[self.from_data_cond & self.to_data_cond & contrast
                                 & department & self.dms_cond].shape[0]
        pay = self.perinatal_mri[self.from_data_cond & self.to_data_cond & contrast
                                 & department & self.pay_cond].shape[0]
        oms = self.perinatal_mri[self.from_data_cond & self.to_data_cond & contrast
                                 & department & self.oms_cond].shape[0]
        pay_and_oms = pay + oms + dms

        dict_cond = {
            "Всего": pay_and_oms,
            "ОМС": oms,
            "Платно": pay + dms,
            "ДМС": dms}

        return dict_cond

    def quarter_report(self) -> dict:
        """
        Search in DB studies with filter and create diction with studies and count of it
        :return: diction with statistic information about quarter studies
        """
        diction = {
            "amb": {
                "with contrast": {"ОМС": {},
                                  "Платно": {},
                                  "ДМС": {},
                                  "Всего": {}},
                "without contrast": {"ОМС": {},
                                     "Платно": {},
                                     "ДМС": {},
                                     "Всего": {}}
            },
            "inpatient": {
                "with contrast": {"ОМС": {},
                                  "Платно": {},
                                  "ДМС": {},
                                  "Всего": {}},
                "without contrast": {"ОМС": {},
                                     "Платно": {},
                                     "ДМС": {},
                                     "Всего": {}}
            }
        }
        diction["amb"].update({"with contrast": self.count_quarter_studies(self.amb_cond, self.contrast_cond)})
        diction["amb"].update({"without contrast": self.count_quarter_studies(self.amb_cond, self.not_contrast_cond)})
        diction["inpatient"].update(
            {"with contrast": self.count_quarter_studies(self.not_amb_cond, self.contrast_cond)})
        diction["inpatient"].update(
            {"without contrast": self.count_quarter_studies(self.not_amb_cond, self.not_contrast_cond)})
        return diction

    def create_year_diction(self) -> dict:
        """
        create new diction for year report
        :return: diction
        """
        diction = {
            f"{dep}\n{cat}": {
                stud: 0 for stud in self.exist_studies()
            } for dep in self.departments_mri['Отделение'] for cat in self.department_list
        }
        for depart in self.not_department_list:
            diction[depart] = {stud: 0 for stud in self.exist_studies()}
            diction[depart] = {'Контраст': 0}
        return diction

    def exam_studies_in_department(self) -> dict:
        """
        if count of studies in department is 0, then delete this department from dictionary
        :return: edited dictionary
        """
        diction = self.create_year_diction()
        for dep in self.departments_mri['Отделение']:
            for cat in self.department_list:
                studies_in_department = self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & self.cat_cond(cat) &
                    self.free_pay_cond & self.cat_cond(cat) & self.dep_cond(dep)].shape[0]
                if studies_in_department == 0:
                    del diction[f"{dep}\n{cat}"]
        return diction

    def year_report(self) -> dict:
        """
        saving count of studies into dictionary for all departments
        :return: diction
        """
        diction = self.exam_studies_in_department()
        for study in self.exist_studies():
            study_cond = self.perinatal_mri['Услуга'] == study
            for dep in self.departments_mri['Отделение']:
                self.add_department_studies(diction, dep, study, study_cond)
            self.add_not_department_studies(diction, study, study_cond)
        return diction

    def add_department_studies(self, diction, dep, study, study_cond):
        """
        saving count of studies into dictionary for department_list
        :param diction: diction
        :param dep: current department
        :param study: current study
        :param study_cond: current study condition
        :return:
        """
        for cat in self.department_list:
            studies = self.perinatal_mri[
                self.from_data_cond & self.to_data_cond & self.dep_cond(dep)
                & self.free_pay_cond & self.cat_cond(cat)].shape[0]
            if studies != 0:
                self.add_cat(diction, dep, study, study_cond, cat)

    def add_not_department_studies(self, diction, study, study_cond):
        """
        saving count of studies into dictionary for studies which not in department
        :param diction: diction
        :param study: current study
        :param study_cond: current study condition
        :return: None
        """
        for cat in self.not_department_list:
            if cat == "Всего":
                studies = self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & study_cond].shape[0]
                diction["Всего"].update({f'{study}': studies})
            else:
                studies = self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & study_cond & self.cat_cond(cat)].shape[0]
                diction[f"{cat}"].update({f'{study}': studies})

                contrast_studies = self.perinatal_mri[
                    self.from_data_cond & self.to_data_cond & self.cat_cond(cat) & self.contrast_cond].shape[0]
                diction[f"{cat}"].update({'Контраст': contrast_studies})

    def add_cat(self, diction: dict, dep: str, study: str, study_cond, cat: str) -> dict:
        """
        saving count of studies into dictionary for all departments
        :param diction: diction
        :param dep: current department
        :param study: current study
        :param study_cond: current study condition
        :param cat: current category condition
        :return: diction
        """
        studies = self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & study_cond & self.dep_cond(dep)
            & self.free_pay_cond & self.cat_cond(cat)].shape[0]
        diction[f"{dep}\n{cat}"].update({f"{study}": studies})

        contrast_study = self.perinatal_mri[
            self.from_data_cond & self.to_data_cond & self.contrast_cond & self.dep_cond(dep)
            & self.free_pay_cond & self.cat_cond(cat)].shape[0]
        diction[f"{dep}\n{cat}"].update({"Контраст": contrast_study})
        return diction
