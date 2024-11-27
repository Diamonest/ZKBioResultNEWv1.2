import tkinter as tk
from tkinter import filedialog
import sqlite3
import pandas as pd

def open_excel_file():
    global filename
    drop_table()
    create_table()
    conn = sqlite3.connect("db.db")
    cursor = conn.cursor()
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Выберите файл Excel",
                                          filetypes=(("Excel files","*.xls"), ("all files", "*.*"))
                                          )

    if filename:
        try:
            df = pd.read_excel(filename, skiprows=1)

            # Объединение имени и фамилии
            df['ФИО'] = df['Фамилия'] + ' ' + df['Имя']  # Предполагается, что у вас есть колонки 'name' и 'surname'

            # Удаление лишних колонок (name и surname) если они не нужны в базе данных.
            df = df.drop(columns=['Имя', 'Фамилия'], errors='ignore') # errors='ignore' - игнорирует ошибки если столбцов нет

            # Запись в базу данных (изменен запрос)
            for row in df.itertuples(index=False):
                cursor.execute("INSERT INTO events (id_events,time,zone_name,device_name,event_dot,event_description,id_employee,ФИО,card_number,id_dep,name_dep,reader_name,test_mode) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", row)
            conn.commit()

            # Остальной код для обработки данных и создания отчета остается тем же, но с учетом измененного имени колонки
            cursor.execute('''
                UPDATE events
                SET id_employee = REPLACE(id_employee, '.0','')
                WHERE id_employee LIKE '%.0'
            ''')

            # ... (остальной код для обработки данных, создание временной таблицы и удаление дубликатов остаётся тем же)

            cursor.execute('''
                SELECT id_events, time, zone_name, device_name, event_dot, event_description, id_employee, ФИО, card_number, id_dep, name_dep, reader_name, test_mode, COUNT(*) AS sum
                FROM events
                GROUP BY id_employee
                ORDER BY COUNT(*) DESC
            ''') # Убрали temp.double, теперь запрос из основной таблицы


            result = cursor.fetchall()
            df_result = pd.DataFrame(result, columns=["id события","Время","Название зоны","Имя устройства","Точка события","Описание события","ID сотрудника","Номер карты","Отдел №","Имя отдела","Имя считывателя","Режим проверки","ФИО","Количество"])
            writer = pd.ExcelWriter('ОтчетЗаМесяц.xlsx', engine='xlsxwriter')
            df_result.to_excel(writer, sheet_name="Sheet1", index = False)
            writer._save()
            conn.close()
            text = tk.Label(root, text="Выполнено")
            text.place(x = 260, y = 210)

        except FileNotFoundError:
            print("Файл не найден.")
        except Exception as e:
            print(f"Произошла ошибка: {e}")

def create_table():
    conn = sqlite3.connect("db.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS events (id_events TEXT, time DATE, zone_name TEXT, device_name TEXT, event_dot TEXT, event_description TEXT, id_employee TEXT, ФИО TEXT, card_number TEXT, id_dep INTEGER, name_dep TEXT, reader_name TEXT, test_mode TEXT)
    ''') # Изменено имя колонки
    conn.commit()
    conn.close()

def drop_table():
    conn = sqlite3.connect("db.db")
    cursor = conn.cursor()
    cursor.execute('''
        DROP TABLE IF EXISTS events
    ''')
    conn.commit()
    conn.close()

root = tk.Tk()
root.title("Сколько обедов за период по человеку")
root.geometry("600x200")
executeRequest = tk.Button(root, text="Выбрать файл", command=open_excel_file)
executeRequest.place(x=260,y=140)

#WITH CTE AS (
#                SELECT [time], id_employee,
#                    ROW_NUMBER() OVER (PARTITION BY DATE([time]), id_employee ORDER BY [time]) AS rn
#                FROM temp.double
#            )
#            DELETE FROM temp.double
#            WHERE ([time], id_employee) IN (
#                SELECT [time], id_employee
#                FROM CTE
#                WHERE rn > 1
#            );

root.mainloop()