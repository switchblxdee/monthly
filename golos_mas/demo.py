"""Create a reproducible, entirely fictional end-to-end example."""
import argparse
import json
from pathlib import Path
from openpyxl import Workbook

T1='Медленный запуск рабочей среды после обновления'
T2='Ошибки подключения к сервису B'
T3='Неудобная навигация в настройках сервиса B'
T4='Недоступность учебного региона X'
T4_DETAIL='Сбой сервиса D в учебном регионе X'


def create_demo(output_dir):
    root=Path(output_dir)
    # Never replace an existing directory, even if it happens to be empty.
    root.mkdir(parents=True,exist_ok=False)
    w=Workbook();current=w.active;current.title='приоритет ноябрь'
    header=['владелец','продукт','тема','~ MAU','кол-во сигналов','сигналы','длительность','влияние','общий балл']
    current.append(header)
    current.append(['SberUser','Рабочая среда A',T1,25000,80,4,2,3,9])
    current.append(['SberWorks','Сервис B',T2,5000,50,4,1,3,8])
    current.append(['SberWorks','Сервис B',T3,5000,1,1,1,2,'UX'])
    current.append(['Прочее','Регион X',T4,25000,75,4,2,3,9])
    previous=w.create_sheet('приоритет октябрь');previous.append(header)
    previous.append(['SberUser','Рабочая среда A',T1,25000,20,1,1,3,5])
    previous.append(['Прочее','Регион X',T4,25000,55,2,1,3,6])
    detail_header=['дата','источник','продукт','тема','Маскированный текст из источника','лайки','сигнал']
    su=w.create_sheet('SU');su.append(detail_header)
    su.append(['03.11.2026','Учебный канал A','Рабочая среда A',T1,'После обновления запуск занимает несколько минут.',44,45])
    su.append(['04.11.2026','Учебный канал A','Рабочая среда A',T1,'Коллеги тоже заметили задержку при запуске.',34,35])
    su.append(['25.10.2026','Учебный канал A','Рабочая среда A',T1,'Пример за предыдущий месяц.',0,1])
    sw=w.create_sheet('SW');sw.append(detail_header)
    sw.append(['10.11.2026','Учебный канал B','Сервис B',T2,'Подключение завершается ошибкой при входе.',49,50])
    sw.append(['11.11.2026','Учебный канал UX','Сервис B',T3,'Нужная настройка находится глубоко в меню.',0,1])
    other=w.create_sheet('прочее');other.append(detail_header)
    other.append(['12.11.2026','Учебный канал D','Сервис D',T4_DETAIL,'Тестовый стенд в регионе X недоступен.',74,75])
    w.save(root/'demo.xlsx')
    statuses=[{'product':'Рабочая среда A','topic':T1,'recurrence':'two','controls':['backlog'],'mitigation':'release','version':'2.0','deadline':'1Q 2027'},
              {'product':'Сервис B','topic':T3,'recurrence':'new'}]
    mapping=[{'priority_product':'Регион X','priority_topic':T4,'detail_product':'Сервис D','detail_topic':T4_DETAIL}]
    promises=[{'id':'DEMO-1','owner':'SberUser','description':'Выпустить учебное обновление','source':'Учебная встреча','initial_due':'2026-11-30','current_due':'2026-12-15','outcome':'postponed','in_work_start':True,'in_work_current':True,'comment':'Учебный пример переноса'},
              {'id':'DEMO-2','owner':'SberWorks','description':'Проверить учебную настройку','source':'Учебная встреча','initial_due':'2026-11-20','outcome':'completed_on_time','in_work_start':True,'in_work_current':False}]
    for name,data in [('statuses.json',statuses),('topic-map.json',mapping),('promises.json',promises)]:
        (root/name).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
    expected={'period':'2026-11','priority_topics':4,'detail_rows_after_filter':5,'sum_signals':206,'promise_due':2,'promise_completed':1,'promise_postponed':1}
    (root/'expected.json').write_text(json.dumps(expected,ensure_ascii=False,indent=2),encoding='utf8')
    return root


def main():
    parser=argparse.ArgumentParser(description='Создать вымышленный набор для проверки Голос IT')
    parser.add_argument('--output-dir',required=True,help='Новая папка для вымышленных данных')
    args=parser.parse_args()
    try: root=create_demo(args.output_dir)
    except FileExistsError: parser.error('Папка уже существует: выберите новое имя')
    print(f'Вымышленные данные созданы: {root}')

if __name__=='__main__': main()
