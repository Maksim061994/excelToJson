from excelToJson import ProcessorExcelToJson
import sys
import os.path


def main(pathLoadExcel, pathSaveJson):
	"""
	Основная функция исполнения кода
		Аргументы:
			- pathLoadExcel - путь, где сохранённ Excel файл
			- pathSaveJson - куда необходимо сохранить json
	"""
	path, extension  = pathLoadExcel.rsplit('.',1)
	if pathSaveJson is None:
		pathSaveJson = path + ".json"
	# Проверка на существование файла
	if os.path.isfile(pathSaveJson):
		print("Заданный выходной файл уже существует. Операция не может быть выполнена.")
		return
	# Загружаем класс
	loader = ProcessorExcelToJson(pathLoadExcel, pathSaveJson, "./settings.ini")
	loader.runProcess()
	return

if __name__ == "__main__":
	pathLoadExcel = sys.argv[1] # чтение первого аргумента командной строки
	pathSaveJson = None
	if len(sys.argv) == 3:
		pathSaveJson = sys.argv[2] # чтение второго аргумента командной строки
	main(pathLoadExcel, pathSaveJson)
