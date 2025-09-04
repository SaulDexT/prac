#подгрузка библиотек и данных
import random
from collections import defaultdict
import openpyxl

#работа с excel
data = openpyxl.Workbook()
sheet = data.active

#создание групп с помощью ввода пользователя
groups = []
num_groups = int(input('Введите количество групп: '))
for x in range(num_groups):
    groups.append('Group_' + str(x+1))

#создание словаря с предметами и учителями с помощью ввода пользователя
less = {}
teachers = []
num_teachers = int(input('Введите количество учителей: '))
for x in range(num_teachers):
    name = input('Введите имя учителя: ')
    lesson = input('Введите предмет: ')
    less[lesson] = name
    if name not in teachers:
        teachers.append(name)

#создание списка кабинетов с помощью ввода пользователя
cabs = []
a = 0
while a != 'stop':
    a = input('Введите номер кабинета (или "stop" для завершения): ')
    if a != 'stop':
        cabs.append(a)

#для упрощения тестирования можно использовать закомментированные данные
'''groups = ['Group_1', 'Group_2', 'Group_3', 'Group_4', 'Group_5']
teachers = ['Ivanov', 'Petrov', 'Sobolev', 'Nikolaev', 'Semenov', 'Ershov']
less = {'Math': 'Ivanov', 'Eng': 'Petrov', 'Inf': 'Sobolev', 'PE': 'Nikolaev', 'Bio': 'Semenov', 'Geo': 'Ershov'}
cabs = ['200', '202', '203', '204', '205', '206']'''

#создание слотов для расписания 
time_slots = [f'Day {a+1} Pair {b+1}' for a in range(6) for b in range(4)]

#защита от задвоения расписания
used_slots = {
    'groups': defaultdict(set),
    'teachers': defaultdict(set),
    'cabs': defaultdict(set)
}

schedule = []

#распределение пар по свободным временным слотам
for les, teacher in less.items():
    for group in groups:
        placed = False
        random.shuffle(time_slots)
        for slot in time_slots:
            if (slot not in used_slots['groups'][group] and
                slot not in used_slots['teachers'][teacher]):
                
                for cab in cabs:
                    if slot not in used_slots['cabs'][cab]:
                        schedule.append({
                            'group': group,
                            'les': les,
                            'teacher': teacher,
                            'cab': cab,
                            'time': slot
                        })
                        used_slots['groups'][group].add(slot)
                        used_slots['teachers'][teacher].add(slot)
                        used_slots['cabs'][cab].add(slot)
                        placed = True
                        break
            if placed:
                break

#запись расписания в excel файл
sheet['A'+str(1)] = ('Группа')
sheet['B'+str(1)] = ('Предмет')
sheet['C'+str(1)] = ('Преподаватель')
sheet['D'+str(1)] = ('Кабинет')
sheet['E'+str(1)] = ('День')
sheet['F'+str(1)] = ('Пара')

for i in range(len(schedule)):
    sheet['A'+str(i+2)] = str(schedule[i]['group'])
    sheet['B'+str(i+2)] = str(schedule[i]['les'])
    sheet['C'+str(i+2)] = str(schedule[i]['teacher'])
    sheet['D'+str(i+2)] = str(schedule[i]['cab'])
    if int((schedule[i]['time'])[4]) == 1:
        sheet['E'+str(i+2)] = 'Пн'
    elif int((schedule[i]['time'])[4]) == 2:
        sheet['E'+str(i+2)] = 'Вт'
    elif int((schedule[i]['time'])[4]) == 3:
        sheet['E'+str(i+2)] = 'Ср'
    elif int((schedule[i]['time'])[4]) == 4:
        sheet['E'+str(i+2)] = 'Чт'
    elif int((schedule[i]['time'])[4]) == 5:
        sheet['E'+str(i+2)] = 'Пт'
    elif int((schedule[i]['time'])[4]) == 6:
        sheet['E'+str(i+2)] = 'Сб'
    sheet['F'+str(i+2)] = str(schedule[i]['time'])[11]
data.save("actual_schedule.xlsx")

print('Ваше расписание создано, и сохранено в той же директории, в которой хранится исполнительный файл в формате .xlsx')