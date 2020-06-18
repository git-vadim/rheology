from math import sqrt, exp
import matplotlib.pyplot as plt
from numpy import array
import pandas as pd
from PIL import ImageTk, Image
from sympy.solvers import solve
from sympy import Symbol
from threading import Thread
from tkinter import Tk, Button, Entry, Label, messagebox, Toplevel
from tkinter.filedialog import askopenfile as fd
import xlwings as xw


# Стек исходных данные
filename = ''
y = []
tay = []
y_X, tay_X, y_XX, tay_XX = 0, 0, 0, 0
mu = 0
mu_0 = 0
tay_0 = 0

# Стек решений
a1_list = []
A1_list = []
a2_list = []
A2_list = []

approx_1_items = []
approx_2_items = []
approx_3_items = []

deltas = []


def get_file():
    global filename
    try:
        filename = fd().name
        file_text.delete(0, 'end')
        file_text.insert(0, filename)
        return filename
    except Exception as error:
        messagebox.showinfo(
            'Ошибка :(', f'Что-то полшло не так:\n{str(error)} \nПопробуйте снова.')

def get_input_values():
    global y, tay
    y_range = input_y_entry.get()
    tay_range = input_tay_entry.get()
    wb = xw.Book(filename)
    sht = wb.sheets['Лист1']
    try:
        y = sht.range(str(y_range)).value
        tay = sht.range(str(tay_range)).value
        print(y, tay)
    except:
        messagebox.showwarning('Внимание', 'Указанный диапазон ячеек не содержит числовых \
                                значений, либо они заданы некорректно.')

def plotting(**kwargs):
    '''y = значение скорости сдвига из исходных данных'''
    x = y
    for key, value in kwargs.items():
        plt.plot(x, value, label=key)
    plt.xlabel('Градиент сдвига, c-1')
    plt.ylabel('Напряжение сдвига, Па')
    plt.title('График зависимости напряжения сдвига от градиента сдвига')
    plt.grid()
    plt.legend()
    plt.show()


def get_points():
    global y_X, tay_X, y_XX, tay_XX, mu, mu0, tay_0
    # точка перегиба
    y_X = float(tay_0_x_entry.get())
    tay_X = float(tay_0_y_entry.get())
    # харак. точка
    y_XX = float(tay_x_x_entry.get())
    tay_XX = float(tay_x_y_entry.get())

    tay_0 = tay[0]
    mu0 = (tay[1]-tay[0])/(y[1])
    mu = tay[-1]/y[-1]

    plt.plot(y, tay, label='Эксперимент')
    plt.plot(y_X, tay_X, 'ro', label='точка перегиба')
    plt.plot(y_XX, tay_XX, 'r+', label='харак.точка')
    plt.xlabel('Скорость сдвига')
    plt.ylabel('Напряжение сдвига')
    plt.title('График')
    plt.grid()
    plt.legend()
    plt.show()

def approx_1():
    global a1_list, A1_list, a2_list, A2_list, approx_1_items
    # части уравнения
    D = 2*(tay_XX-tay_0)/(y_XX**2) - 2*mu0/y_XX
    D1 = D/(tay_X-mu*y_X)
    k4 = y_X*(mu-mu0)**2
    k3 = -2*D*(mu-mu0)*y_X + (mu-mu0)*(mu-mu0) + D1*tay_0**2
    k2 = (y_X)*(D**2)-2*D*(mu-mu0)
    k1 = D*D-D1*((mu-mu0)*(mu-mu0) + 2*tay_0*D)
    k0 = D1*(mu-mu0)*2*D
    # решение самого уравнения
    a1 = Symbol('a1')
    root = solve((a1**4)*k4 + (a1**3)*k3 + (a1**2)*k2 + (a1*k1) + k0, a1)
    # избавляемся от комплексной части
    R = [i.as_real_imag() for i in root]
    ROOT = [i[0] for i in R]
    # получаем Корень уравнения
    a1 = max(ROOT)
    A1 = D/(a1**2)
    a2 = (mu-mu0-a1*A1)/(tay_0-A1)
    A2 = tay_0 - A1

    a1_list.append(a1)
    A1_list.append(A1)
    a2_list.append(a2)
    A2_list.append(A2)

    approx_1_items = [A1*exp(-a1*i)+A2*exp(-a2*i)+mu*i for i in y]

    get_delta(approx_1_items)
    messagebox.showinfo(title='Поправочный коэффициент а1',
                        message=f'Получено значение а1={a1}')

def approx_2():
    global a1_list, A1_list, a2_list, A2_list, approx_2_items
    a1 = float(a1_entry.get())
    A1 = A1_list[0]
    A2 = A2_list[0]
    a2 = (mu-mu0-a1*A1)/(tay_0-A1)

    a1_list.append(a1)
    A1_list.append(A1)
    a2_list.append(a2)
    A2_list.append(A2)

    approx_2_items = [A1*exp(-a1*i)+A2*exp(-a2*i)+mu*i for i in y]

    get_delta(approx_2_items)

def approx_3():
    global a1_list, A1_list, a2_list, A2_list, approx_3_items
    a1 = a1_list[1]

    R0 = exp(a1*y_X)*(tay_X-mu*y_X)*(mu-mu0)**2
    A1 = ((2*(mu-mu0)*a1*exp(a1*y_X)*(tay_X-mu*tay_X) -
           (a1**2*tay_0**2-(mu-mu0)**2) -
           sqrt((2*(mu-mu0)*a1*exp(a1*y_X)*(tay_X-mu*y_X) -
                 (a1**2*tay_0**2-(mu-mu0)**2))**2 -
                4*((a1**2*exp(a1*y_X)*(tay_0-mu*y_X) +
                    2*a1*(mu-mu0-a1*tay_0))*R0))) /
          (2*(a1**2*exp(a1*y_X)*(tay_X-mu*y_X)+2
              * a1*(mu-mu0-a1*tay_0))))
    a2 = (mu-mu0-a1*A1)/(tay_0-A1)
    A2 = tay_0-A1

    a1_list.append(a1)
    A1_list.append(A1)
    a2_list.append(a2)
    A2_list.append(A2)

    approx_3_items = [A1*exp(-a1*i)+A2*exp(-a2*i)+mu*i for i in y]

    get_delta(approx_3_items)

def get_delta(x):
    global deltas
    tay_np = array(tay)
    A = (x - tay_np).tolist()
    square_sqrt = [sqrt(i**2) for i in A]
    delta = sum(square_sqrt)/len(tay)
    deltas.append(delta)

def get_table_view():

    df = pd.DataFrame({
        'a1': [round(i, 3) for i in a1_list],
        'A1': [round(i, 3) for i in A1_list],
        'a2': [round(i, 3) for i in a2_list],
        'A2': [round(i, 3) for i in A2_list],
        'delta': [round(i, 3) for i in deltas]
    })
    df.index = ['Approx_1','Approx_2','Approx_3']

    window = Toplevel(root)
    height = 5
    width = 5
    for i in range(height): #Rows
        for j in range(width): #Columns
            for item in df.items():
                b = Label(window, text=str(item))
                b.grid(row=i, column=j)

    return df

def finish():
    # print(get_table_view())
    best = deltas.index(min(deltas))
    eq = f'a1={"%.2f" % a1_list[best]}, \
A1={"%.2f" % A1_list[best]}, \
a2={"%.2f" % a2_list[best]}, \
A2={"%.2f" % A2_list[best]}'
    
    messagebox.showinfo('Решение', f'Уравнение, наилучшим образом описывающее кривую, \
содержит следующие коэффициенты: \n{eq}')

def clear_all():
    global a1_list, A1_list, a2_list, A2_list, approx_1_items, approx_2_items, approx_3_items, deltas
    
    a1_list = []
    A1_list = []
    a2_list = []
    A2_list = []

    approx_1_items = []
    approx_2_items = []
    approx_3_items = []

    deltas = []


# Дезайн
root = Tk()
root.geometry('960x400+300+100')
root.title('Расчёт реологического уравнения')


# Objects
file_button = Button(root, text='Выбрать Excel файл',
                     command=get_file, cursor='pointinghand')
file_text = Entry(root)

img = ImageTk.PhotoImage(Image.open('algorithm.png'))
image_label = Label(root, image=img)
image_label.place(x=400, y=10)

# for input
input_tay_label = Label(root, text='Исх. T:')
input_y_label = Label(root, text='Исх. Y:')

input_tay_entry = Entry(root, width=7)
input_y_entry = Entry(root, width=7)

input_values_button = Button(root, text='Исх. данные',
                             command=get_input_values,
                             width=9)
input_values_plot = Button(
    root, text='График', command=lambda: plotting(Эксперимент=tay))

# char dots
x_label = Label(root, text='Y')
y_label = Label(root, text='TAY')
tay_0_label = Label(root, text='Перег.')
tay_0_x_entry = Entry(root, width=2)
tay_0_y_entry = Entry(root, width=2)

tay_x_label = Label(root, text='Харак.')
tay_x_x_entry = Entry(root, width=2)
tay_x_y_entry = Entry(root, width=2)

dots_button = Button(root, text='Отметить точки', command=get_points)

# Approx 1
approx_1_button = Button(root, text='Первое приближение', command=approx_1)
approx_1_plot = Button(root, text='График', command=lambda: plotting(
    Эксперимент=tay, Approx_1=approx_1_items))
# Approx 2
a1_label = Label(root, text='Поправка а1')
a1_entry = Entry(root, width=5)
approx_2_button = Button(root, text='Второе приближение', command=approx_2)
approx_2_plot = Button(root, text='График', command=lambda: plotting(
    Эксперимент=tay, Approx_1=approx_1_items, Approx_2=approx_2_items))
# APPROX 3
approx_3_button = Button(root, text='Третье приближение', command=approx_3)
approx_3_plot = Button(root, text='График', 
                             command=lambda: plotting(Эксперимент=tay, 
                                                      Approx_1=approx_1_items, 
                                                      Approx_2=approx_2_items, 
                                                      Approx_3=approx_3_items))


final = Button(root, text='FINISH', command=finish)
final.place(x=250, y=300)

# Places
file_button.place(x=10, y=10)
file_text.place(x=150, y=5)

# for input
input_y_label.place(x=10, y=50)
input_y_entry.place(x=10, y=70)

input_tay_label.place(x=110, y=50)
input_tay_entry.place(x=110, y=70)

input_values_button.place(x=190, y=73)
input_values_plot.place(x=280, y=73)

# char dots
x_label.place(x=65, y=105)
y_label.place(x=88, y=105)
tay_0_label.place(x=10, y=128)
tay_0_x_entry.place(x=60, y=125)
tay_0_y_entry.place(x=90, y=125)

tay_x_label.place(x=10, y=155)
tay_x_x_entry.place(x=60, y=152)
tay_x_y_entry.place(x=90, y=152)

dots_button.place(x=130, y=140)

# approx 1
approx_1_button.place(x=10, y=200)
approx_1_plot.place(x=160, y=200)
# approx 2
a1_label.place(x=10, y=250)
a1_entry.place(x=100, y=247)
approx_2_button.place(x=170, y=250)
approx_2_plot.place(x=320, y=250)
# approx 3
approx_3_button.place(x=10, y=300)
approx_3_plot.place(x=160, y=300)
# clear
clear_button = Button(root, text='Clear', command=clear_all)
clear_button.place(x=250, y=350)


if __name__ == '__main__':
    messagebox.showwarning(
        'Внимание', 'Перед использованием программы внимательно \
изучите инструкцию в файле readme.txt')
    root.mainloop()
