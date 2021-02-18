# -*- coding: utf-8 -*-
import glob
import os
import pandas as pd
import numpy as np
import re
from pandas.api.types import is_numeric_dtype


class SuccessAnalyzer:
    """Класс, который реализует считывание excel таблиц и анализ успешности сотрудников"""
    __dict_table = {}
    __united_table = pd.DataFrame()

    def __init__(self, path_read, path_output=None):
        """Инициализация обьекта класса"""
        self.path_read = path_read
        if path_output is None:
            self.path_output = path_read
        else:
            self.path_output = path_output

    def __assessment_success(self, column_plan, column_actual):
        """Высчитываем оценку сотрудника по каждому проекту"""
        if is_numeric_dtype(type(column_plan)) and is_numeric_dtype(type(column_actual)):
            if np.isnan(column_plan) or np.isnan(column_actual):
                return 0
            elif column_plan == 0 and column_actual:
                return 1
            elif column_plan and column_actual == 0:
                return 0
            elif column_plan == column_actual:
                return 1
            elif column_plan < column_actual:
                return 0.5
            elif column_plan > column_actual:
                return 1.5
        else:
            return 0

    def __create_name_column(self, old_column_name):
        """Получаем имя сотрудника из старого имени столбца"""
        try:
            return re.search(r"[А-ЯЁ][а-яё]+\s+[А-ЯЁ]+\.+[А-ЯЁ]+\.?", old_column_name).group(0)
        except (TypeError, AttributeError):
            return "Unnamed"

    def __create_table_success(self, table):
        """Создаем таблицу с анализом успешности сотрудников"""
        new_table = pd.DataFrame()
        for i in range(3, table.shape[1], 2):
            slice_table = table.iloc[:, i:i + 2]
            new_column = slice_table.apply(lambda x: self.__assessment_success(x[0], x[1]), axis=1)
            name_column = self.__create_name_column(slice_table.iloc[:, 0].name)
            new_table[name_column] = new_column
        table_columns_mean = new_table.mean(axis=0).sort_values(ascending=False)
        new_table = pd.DataFrame({"Имя": table_columns_mean.index, "Оценка": table_columns_mean.values})
        return new_table

    def __start_analyzer(self):
        """"Считывание и анализ таблиц с последующим сохранением данных в переменные класса"""
        filename_list = glob.glob("{}*.xlsx".format(self.path_read))
        if filename_list:
            for path_to_read in filename_list:
                excel_table = pd.read_excel(path_to_read, index_col=0)
                table_assessment = self.__create_table_success(table=excel_table)
                name_file = os.path.splitext(os.path.basename(path_to_read))[0]
                self.__united_table = pd.concat([self.__united_table, table_assessment], ignore_index=True)
                self.__dict_table[str(name_file)] = table_assessment
            self.__united_table = self.__united_table.sort_values(by=["Оценка"], ascending=False)

    def print_united_table(self):
        """"Считывание данных, их анализ и вывод общей таблицы по сотрудникам"""
        filename_list = glob.glob("{}*.xlsx".format(self.path_read))
        table = pd.DataFrame()
        name_table = "Анализ успешности сотрудников"
        if filename_list:
            for path_to_read in filename_list:
                excel_table = pd.read_excel(path_to_read, index_col=0)
                table_assessment = self.__create_table_success(table=excel_table)
                table = pd.concat([table, table_assessment], ignore_index=True)
            print("Таблица {0}:\n{1}".format(name_table,
                                             table.groupby(['Имя'], as_index=False).sum().sort_values(by=["Оценка"],
                                                                                                      ascending=False,
                                                                                                      ignore_index=True)))

    def output_to_file(self):
        """"Считывание данных, их анализ и вывод по каждой таблице в файл формата xlsx"""
        filename_list = glob.glob("{}*.xlsx".format(self.path_read))
        if filename_list:
            for path_to_read in filename_list:
                excel_table = pd.read_excel(path_to_read, index_col=0)
                table_assessment = self.__create_table_success(table=excel_table)
                name_file = os.path.splitext(os.path.basename(path_to_read))[0]
                table_assessment.to_excel("{0}{1}_assessment.xlsx".format(self.path_output, name_file), index=False)

    def output_to_console(self):
        """"Считывание данных, их анализ и вывод по каждой таблице в консоль"""
        filename_list = glob.glob("{}*.xlsx".format(self.path_read))
        for path_to_read in filename_list:
            excel_table = pd.read_excel(path_to_read, index_col=0)
            table_assessment = self.__create_table_success(table=excel_table)
            name_file = os.path.splitext(os.path.basename(path_to_read))[0]
            print("Таблица {}:".format(name_file))
            print(table_assessment.head())
