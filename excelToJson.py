import pandas as pd
from pandas.api.types import is_numeric_dtype, is_datetime64_dtype
import numpy as np
import json
# Подключение файла с конифигами сервера
import configparser

class ProcessorExcelToJson:
    """Класс преобразования Excel в Json"""
    def __init__(self, pathToLoadFile, pathToSaveFile, pathConfig="./"):
        """
            pathToLoadFile - путь до Excel файла
            pathToSaveFile - путь для сохранения файла
            listSheets - список sheets
            pathConfig* - путь к конфигу
        """
        # Читаем конфиг
        config = configparser.ConfigParser()
        config.read(pathConfig, encoding="utf-8")
        # Чтения данных из конфига
        self.coding = config.get("SETUP", "CODING") # кодировка
        self.listSheets = config.get("SETUP", "LIST_SHEETS") # список вкладок
        # Путь к файлам для чтения и сохранения
        self.pathToExcel = pathToLoadFile
        self.pathToSaveFile = pathToLoadFile
        # Путь задан, то используется тот путь, которые прописал пользователь
        if pathToSaveFile:
            self.pathToSaveFile = pathToSaveFile
    
    def changeTypeColumn(self, df):
        """
        Метод изменения типов колонок
            Аргументы:
                - df - входной датафрейм
            Результат:
                - df - датафрейм с преобразованными типами
        """
        for col in df.columns:
            # Если тип колонок DateTime - преобразуем в строку
            if is_datetime64_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%dT%H:%M:%S")
                continue
            if is_numeric_dtype(df[col]):
                # df[col] = df[col].astype('Int32')   
                if col == "FINAL_GUILTY_FIRM":
                    df[col] = df[col].fillna(0).astype(int).replace(0, None)
                continue
            df[col] = df[col].astype(str)
        if "Unnamed: 0" in df.columns:
            del df["Unnamed: 0"]
        try:
            df = df.replace("nan", "")
        except Exception as e:
            print("Error with replace nan :", e)
        return df
    
    def procOneFile(self, sheet):
        """
        Чтение и обработка одного датафрейма
            Аргументы:
                - sheet - вкладка, на котором находится датафрейм
            Результат:
                dictDf - dict Python
        """
        # Загружаем датафрейм
        df = pd.read_excel(self.pathToExcel, sheet_name=sheet)
        # Изменяем типы колонок
        df = self.changeTypeColumn(df)
        # NaN в Null
        df = df.where((pd.notnull(df)), None)
        dictDf = df.to_dict(orient="records")
        return dictDf

    
    def runExcelToDict(self):
        """
        Метод получение из файла Excel dict Python
            Аргументы:
                -
            Результат:
                - dictData - словарь Python
        """
        dictData = dict()
        # Получаем список sheets
        xl = pd.ExcelFile(self.pathToExcel)
        sheetNames = xl.sheet_names
        # Читаем данные
        for sheet in sheetNames:
            # Если вкладка не описана в файле settings - не чиатем данную вкладку
            if sheet not in self.listSheets:
                continue
            print("Обрабатываю вкладку - ", sheet)
            dictDf = self.procOneFile(sheet) # получаем массив словарей по одной вкладке
            dictData[sheet] = dictDf
        return dictData

    def actsAddProcessJson(self, items, listChangeValue=["IT_SECTIONS", "IT_SIGNS", "IT_CREW", "IT_INVENT"]):
        """
        Дополнительная обработка для преобразования под Json-ов
            Аругменты:
                - items - массив объектов
                - listChangeValue* - список колонок для преобразования
        """
        for item in items:
            for val in listChangeValue:
                # Если val нет среди ключей - то не преобразуем
                if val not in item.keys():
                    continue
                item[val] = json.loads(item[val])
            for key, value in item.items():
                if type(value) == float:
                    # Обработка float. Нетрализация не нужных преобрзований во float. Пример: если  2.0 - 2 == 0 --> преобразуем в int
                    if value - int(value) == 0:
                        item[key] = int(value)
        return items

    def runProcess(self):
        """
        Преобразование Excel в Json
        """
        dictData = self.runExcelToDict()
        # Обработка внутренних JSON-ов актах
        print("Провожу последние преобразования")
        for key, items in dictData.items():
            # if "acts" not in key:
            #     continue
            items = self.actsAddProcessJson(items)
            dictData[key] = items
        # Сохранение JSON-а
        jsData = json.dumps(dictData, ensure_ascii=False, indent=4).encode(self.coding).decode(self.coding)
        jsData = jsData.replace("\"NaT\"", "null")
        with open(self.pathToSaveFile, "w", encoding=self.coding) as f:
            f.write(jsData)
        