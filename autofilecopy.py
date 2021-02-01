import re , os , shutil , openpyxl

from os import path
base_dir = path.dirname(path.abspath(__file__))
#config_path = path.join(base_dir, 'settings.ini')

#---------------------------------------------------------------------------

def zipper(list_number):	#задаем функцию для работы с excel файлом
	
	wb = openpyxl.load_workbook(str(base_dir)+'/list.xlsx')	#открываем excel лист по пути
	
	wb.active=list_number	#задаем активный лист
	sheet=wb.active	#sheet - номер листа
	
	A1=[]	#пустой список для столбца 'A'
	for cell in sheet['A']:	#цикл для всех значений в столбце 'A'
		A1.append(cell.value)	#добавляем каждый элемент в список А1
		
	B1=[]	##пустой список для столбца 'B'
	for cell in sheet['B']:	#цикл для всех значений в столбце 'B'
		B1.append(cell.value)	#добавляем каждый элемент в список B1

	A1_B1=dict(zip(A1,B1))	#склеиваем списки 'A1' и 'B1' в словарь 'A1_B1'
	return A1_B1	#Возвращаем словарь
	
#---------------------------------------------------------------------------

def copyer(build_num):	#задаем функцию с аргументом номер здания(номер листа в excel)

	pytj = str(base_dir)+'/teplak/'	#путь к папке с фотографиями
	
	list = os.listdir(pytj)	#создаем список с названиями фотографий
	string = '_'.join(list)	#преобразуем этот список в строку с разделителем '_'


#фильтр присваивающий имя(name) в зависимости от номера здания(build_num)
	if build_num == 0:	
		name = 'АВК'
		print(name+' : ')
	elif build_num == 1:
		name = 'АТК'
		print(name+' : ')
	elif build_num == 2:
		name = 'ВДС'
		print(name+' : ')
	elif build_num == 3:
		name = 'ГКО'
		print(name+' : ')
	elif build_num == 4:
		name = 'Грузовой'
		print(name+' : ')
	elif build_num == 5:
		name = 'Котельная'
		print(name+' : ')
	elif build_num == 6:
		name = 'ОМТС'
		print(name+' : ')
	elif build_num == 7:
		name = 'ПБЗ'
		print(name+' : ')
	elif build_num == 8:
		name = 'СРТ'
		print(name+' : ')
	elif build_num == 9:
		name = 'DutyFree'
		print(name+' : ')
	elif build_num == 10:
		name = 'VIP'
		print(name+' : ')

		
	build = zipper(build_num)	#задаем переменную (build) как функцию zipper с аргументом build_num - номер здания(номер листа в excel файле)

	
	for k in build:	#цикл запускающий 'zipper'  k - номер здания, номер листа в excel
		
		excel_string = str(build[k])	#создаем динамическую строку с номерами фотографий в столбце 'B1' excel файла
		
		excel_string = excel_string.replace(' ','.')	#заменяем разделитель 'пробел' на точку для удобства
		
		excel_key = excel_string.split('.')
			#преобразовываем эту строку в динамический список
		
		i = len(excel_key)	#количесво фотографий в одну папку
		
		for i in range(0,i):	#цикл перебирающий каждый номер фото в ключ( i - либо ноль фото если пусто либо сколько записано в excel файле )
		
			key = str(excel_key[i])+'.BMT'	#формируем имя файла добавляя к его имени расширение как строку
			
			sum = len(key)	#sum - количество символов в имени (пример '26.BMT' (6 символов))

#В зависимости от суммы символов в имени так как конечное количество символов всегда одинаковое ( IR1234567.BMT - 12 ) подставляем в начало имени 'IR' и определенное количество нолей зависящее от суммы символов определенных ранее


			if sum == 5:
				IR= 'IR00000'
			elif sum == 6:
				IR= 'IR0000'
			elif sum == 7:
				IR= 'IR000'
			elif sum == 8:
				IR= 'IR00'
			elif sum == 9:
				IR= 'IR0'
			elif sum == 10:
				IR= 'IR'
				
			
			match = re.search(key,string)	#задаем переменную match(совпадение) как функцию поиска совпадений между 'key' (нашем сформированном имени) и 'string'(строка сформированная из списка файлов в папке 'Teplak')
		
			if match:	#если совпало, то:
			
				file = str(base_dir)+'/Teplak/'+IR+key+''	#задаем переменную file как локальный путь к файлу на устройстве подствляя части имени (IR = ('IR000') и key = ('228.BMT'))
				
				papka = str(base_dir)+'/Tree/'+name+'/'+str(k)+''	#задаем переменную papka как локальный путь к папке на устройстве подствляя названия нужных папок (name = ('АВК') и str(k) = ('01ЩАВК'))
				
				copy=shutil.copy(file,papka)	#задаем переменную 'copy' как функцию модудя 'shutil' с аргументами 'file' - локальный путь к файлу, 'papka' - локальный путь к папке, функция копирует файл(file) в папку(papka)

#---------------------------------------------------------------------------

for i in range(0,10):	#задаем цикл от i до 10 (по очереди будет прогонять каждое здание)
	
	copyer(i)	#(переменная 'i' будет поочередно меняться от '0' до '10' , каждая смена происходит после завершения функции finder)
	print('Готово!')	#выводим 'Готово!' после каждого цикла одного значения переменной 'i'

