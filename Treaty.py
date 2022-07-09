import sys
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Combobox
from docxtpl import DocxTemplate
    
def help_0():
    txt_1 = Label(frame, text='Номер/месяц/год: 128 06 22.', bg='white', fg='black')
    txt_1.place(x=570, y=5)

def help_1():
    txt_1 = Label(frame, text='Пример: 31 октября 2022г.', bg='white', fg='black')
    txt_1.place(x=570, y=37)

def help_2():
    txt_2 = Label(frame, text='Серия и номер: 0123 456789.', bg='white', fg='black')
    txt_2.place(x=570, y=107)

def help_3():
    txt_3 = Label(frame, text='Ваше место проживания где вы прописаны:\nПример:Регион,  г., ул., д., кв. ', bg='white', fg='black')
    txt_3.place(x=570, y=143)

def help_4():
    txt_4 = Label(frame, text='Пример: +7 (997) 183-32-45', bg='white', fg='black')
    txt_4.place(x=570, y=177)

def help_5():
    txt_5 = Label(frame, text='Пример: Регион, город, ул. и т.д.', bg='white', fg='black')
    txt_5.place(x=570, y=248)
def help_6():
    txt_6 = Label(frame, text='Опишите объект. Пример: \nтрёхкомнатная квартира общей площадью 60кв. м.', bg='white', fg='black')
    txt_6.place(x=570, y=283)

def exit_programm():
    sys.exit()

def files():
    
    #получение с ввода
    n1 = num1.get()
    n2 = num2.get()
    n3 = num3.get()
    d = date.get()
    n = name.get()
    nums = number.get()
    adr = adress.get()
    ph = phone.get()
    obj = object1.get()
    loca = location.get()
    dscr = discription.get()
    val_type = value_type.get()
    obj_type = object_type.get()
    dat = date1.get()
    val_price = price.get()
    term = time1.get()

    if n1 == '' or n2 == '' or n3 == '' or d == '' or n == '' or nums == '' or adr == '' or ph == '' or obj == '' or loca == '' or dscr == '' or val_type == '' or obj_type == '' or dat == '' or val_price == '' or term == '':
        messagebox.showinfo('Error', 'Все поля обязательны для заполнения!')
        return

    else:   
        doc = DocxTemplate('Шаблон.docx')
        context = {'num': n1, 'month': n2, 'year': n3, 'dateofcontract': d, 'fio': n, 'passport': nums, 'address': adr, 'phone': ph, 'typeobject': obj, 'addressobject': loca, 'objectdescription': dscr, 'typecost': val_type,  'evaluationpurpose': obj_type, 'datecost': dat, 'cost': val_price, 'ter': term}
        doc.render(context)
        doc.save('Dogovor_KOREL.docx')
        messagebox.showinfo('...', 'Файл с отчётом был успешно создан.')
        exit_programm()

window = Tk()
window.title('Договор на оценку')
window['bg'] = '#307DEE'
window.geometry('940x680+200+5')
window.resizable(width=False, height=False)

frame = Frame(window, bg='white')
frame.place(relx=0.005, rely=0.005, relwidth=0.99, relheight=0.99)


def erase():
    num1.delete(0, 'end')
    num2.delete(0, 'end')
    num3.delete(0, 'end')
    date.delete(0, 'end')
    name.delete(0, 'end')
    number.delete(0, 'end')
    adress.delete(0, 'end')
    phone.delete(0, 'end')
    object1.delete(0, 'end')
    location.delete(0, 'end')
    discription.delete(0, 'end')
    value_type.delete(0, 'end')
    object_type.delete(0, 'end')
    date1.delete(0, 'end')
    price.delete(0, 'end')
    time1.delete(0, 'end')
 
#номер договора
txt1 = Label(frame, text="№ договора: ", bg='#4C90E8', width=30)
txt1.place(x=10, y=6)
num1 = Entry(frame, width=4, bg='white')
num2 = Entry(frame, width=4, bg='white')
num3 = Entry(frame, width=4, bg='white')
num1.place(x=230, y=7)
num2.place(x=260, y=7)
num3.place(x=290, y=7)

#дата
txt2 = Label(frame, text="Дата заключения договора: ", bg='#4C90E8', width=30)
txt2.place(x=10, y=40)
date = Entry(frame, bg='white', width=50)
date.place(x=230, y=41)

#фио
txt3 = Label(frame, text="Ф. И. О. клиента: ", bg='#4C90E8', width=30)
txt3.place(x=10, y=75)
name = Entry(frame, bg='white', width=50)
name.place(x=230 , y=76)

#паспортные данные
txt4 = Label(frame, text="Паспортные данные: ", bg='#4C90E8', width=30)
txt4.place(x=10, y=110)
number = Entry(frame, bg='white', width=50)
number.place(x=230, y=111)

#адрес
txt5 = Label(frame, text="Адрес регистрации: ", bg='#4C90E8', width=30)
txt5.place(x=10, y=145)
adress = Entry(frame, bg='white', width=50)
adress.place(x=230, y=146)

#номер телефона
txt6 = Label(frame, text="Номер телефона клиента: ", bg='#4C90E8', width=30)
txt6.place(x=10, y=180)
phone = Entry(frame, bg='white', width=50)
phone.place(x=230, y=180)

#oбъект
txt7 = Label(frame, text="Вид объекта оценки: ", bg='#4C90E8', width=30)
txt7.place(x=10, y=215)
object1 = Combobox(frame, width=47)  
object1['values'] = ("Квартира", "Дом", "Земельный участок", "Автомобиль")
object1.place(x=230, y=215)

#Местонахождение
txt8 = Label(frame, text="Местонахождение объекта оценки: ", bg='#4C90E8', width=30)
txt8.place(x=10, y=250)
location = Entry(frame, bg='white', width=50)
location.place(x=230, y=251)

#Описание
txt9 = Label(frame, text="Описание объекта оценки: ", bg='#4C90E8', width=30)
txt9.place(x=10, y=285)
discription = Entry(frame, bg='white', width=50)
discription.place(x=230, y=286)

#определить вид стоимости
txt10 = Label(frame, text="Вид определяемой стоимости: ", bg='#4C90E8', width=30)
txt10.place(x=10, y=320)
value_type = Combobox(frame, width=47)  
value_type['values'] = ("Рыночная", "Рыночная и ликвидационная")
value_type.place(x=230, y=320)

#цель
txt11 = Label(frame, text="Цель оценки: ", bg='#4C90E8', width=30)
txt11.place(x=10, y=355)
object_type = Combobox(frame, width=47)  
object_type['values'] = ("Определение рыночной стоимости", "Определение рыночной и ликвидационной стоимости")
object_type.place(x=230, y=355)

#дата определения
txt12 = Label(frame, text="Дата определения объекта стоимости: ", bg='#4C90E8', width=30)
txt12.place(x=10, y=390)
date1 = Entry(frame, bg='white', width=50)
date1.place(x=230, y=391)

#cтоимость услуг
txt13 = Label(frame, text="Cтоимость услуг: ", bg='#4C90E8', width=30)
txt13.place(x=10, y=425)
price = Entry(frame, bg='white', width=50)
price.place(x=230, y=426)

#срок подготовки
txt14 = Label(frame, text="Срок подготовки отчёта об оценки: ", bg='#4C90E8', width=30)
txt14.place(x=10, y=460)
time1 = Entry(frame, bg='white', width=50)
time1.place(x=230, y=461)

#выход
exit1 = Button(frame, text='Закрыть', bg='#4C90E8', fg='white', command=exit_programm)
exit1.place(x=8, y=638)

#сформировать договор
a = Button(frame, text='Cформировать \nдоговор', bg='#4C90E8', fg='white', command=files)
a.place(x=822, y=628)

error = Label(frame, text="")
error.place(x=700, y=600)

#кнопки help
btn1 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_0)
btn1.place(x=538, y=5)
btn2 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_1)
btn2.place(x=538, y=37)
btn4 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_2)
btn4.place(x=538, y=107)
btn5 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_3)
btn5.place(x=538, y=143)
btn6 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_4)
btn6.place(x=538, y=177)
btn8 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_5)
btn8.place(x=538, y=248)
btn8 = Button(frame, text="?", bg='#4C90E8', width=2, fg='white',command=help_6)
btn8.place(x=538, y=283)
btn_erase = Button(frame, text='Cброс', bg='#4C90E8', fg='white',command=erase)
btn_erase.place(x=10, y=500)
style = Label(frame, text='------------------------------------------', bg='white')
style.place(x=318, y=6)
lbl = Label(frame, bg='white', text='______________________________________________________________________________________________________________________________________________________________________________________________________')
lbl.place(x=-10, y=608)
window.mainloop()
