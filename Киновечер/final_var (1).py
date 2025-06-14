# -*- coding: utf-8 -*-
"""
Created on Sat May 13 14:41:02 2023

@author: Никита, Тимур и Денис
"""
import os
import tkinter as tki
import tkinter.ttk as ttk
from tkinter import messagebox
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns

os.chdir("D:/Work/Data")
cwd = os.getcwd()
os.listdir('.')
#                                                           ОТСЮДОВА
file = pd.read_excel('moviedataset.xlsx')  # ЕСЛИ УЖЕ ЕСТЬ ГЛОБАЛЬНАЯ ПЕРЕМЕННАЯ ПОД ФАЙЛ, ТО ПЕРЕОБОЗНАЧИТЬ

var_1 = None
var_2 = None
var_3 = None
var_4 = None
var_5 = None
tree = None


def search_film():
    window_search_var1 = tki.Toplevel(root)
    window_search_var1.wm_attributes("-topmost", 1)
    window_search_var1.title("Search film")
    window_search_var1.geometry('700x550+50+50')
    window_search_var1.resizable(width=0, height=0)
    window_search_var1["bg"] = "gray23"
    logo_text = tki.Label(window_search_var1, text="Подборка фильмов", foreground="#33CCFF", background="gray23",
                          font=("Gilroy ExtraBold", 30))
    logo_text.place(relx=0.25, rely=0.4)
    textforganres = tki.Label(window_search_var1, text="Жанр:", foreground="white", background="gray23",
                              font=("Gilroy ExtraBold", 14))
    textforganres.place(relx=0.08, rely=0.13)
    textforganres = tki.Label(window_search_var1, text="Хронометраж \n в мин. :", foreground="white",
                              background="gray23", font=("Gilroy ExtraBold", 14))
    textforganres.place(relx=0.23, rely=0.11)
    textforganres = tki.Label(window_search_var1, text="Финансируемость \n фильма:", foreground="white",
                              background="gray23", font=("Gilroy ExtraBold", 14))
    textforganres.place(relx=0.45, rely=0.11)
    textforganres = tki.Label(window_search_var1, text="Страна:", foreground="white", background="gray23",
                              font=("Gilroy ExtraBold", 14))
    textforganres.place(relx=0.72, rely=0.13)
    textforganres = tki.Label(window_search_var1, text="Язык: ", foreground="white", background="gray23",
                              font=("Gilroy ExtraBold", 14))
    textforganres.place(relx=0.85, rely=0.13)
    # textforganres = tki.Label(window_search_var1, text = "Жанр", foreground="#33CCFF", background="gray23", font=("Gilroy ExtraBold",17) )
    # textforganres.place(relx=0.08, rely=0.15)
    under_logo_text = tki.Label(window_search_var1, text="Найдите фильмы для себя в любое время суток",
                                foreground="white",
                                background="gray23", font=("Gilroy ExtraBold", 10))
    under_logo_text.place(relx=0.305, rely=0.5)
    main_menu = tki.Menu(window_search_var1, tearoff=0, foreground="#33CCFF", background="gray21", relief="ridge",
                         borderwidth=2)
    jenre_list = ['Comedy', 'Animation', 'Horror', 'Fantasy', 'Thriller', 'Action', 'Romance', 'Mystery', 'Adventure',
                  'Drama', 'Western', 'Crime', 'Family']
    time_list = ['до 60', 'от 60 до 90', 'от 90 до 120', 'от 120 до 150', 'от 150']
    budget_list = ['Арт-хаус', 'Низкобюджетное', 'Среднебюджетное', 'Высокобюджетное', 'Блокбастер']
    language_list = ['Любой', 'en', 'de', 'he', 'es', 'zh', 'ja', 'fr', 'da', 'it', 'sv', 'no', 'pt', 'nl', 'ko', 'af',
                     'fi',
                     'ro', 'ru', 'hi', 'hu', 'bs', 'th', 'cn', 'el', 'pl', 'sr', 'fa', 'ps', 'cs', 'bo', 'bg', 'ta',
                     'tr', 'te', 'sl', 'la', 'nb', 'ca', 'xx', 'bm', 'id', 'ml', 'lv', 'lo', 'ar', 'is', 'kn', 'ku',
                     'mr', 'et', 'ur', 'sq']
    country_list = ['Любая', 'US', 'AR', 'FR', 'DE', 'IL', 'GB', 'MX', 'ES', 'AT', 'CN', 'JP', 'UY', 'IE', 'AU', 'NZ',
                    'CA',
                    'IT', 'CZ', 'DK', 'FI', 'IN', 'BR', 'BE', 'BS', 'KR', 'ZA', 'NL', 'CH', 'HK', 'MY', 'BG', 'DZ',
                    'RU', 'IS', 'PE', 'TW', 'HU', 'TH', 'LU', 'RO', 'RS', 'CL', 'SG', 'NA', 'NO', 'IR', 'SE', 'PL',
                    'AW', 'JM', 'MA', 'AF', 'GR', 'EC', 'AE', 'LY', 'TR', 'HR', 'SI', 'LT', 'EG', 'PT', 'BY', 'UA',
                    'ML', 'BF', 'KZ', 'PH', 'ID', 'LV', 'BA', 'EE', 'GE', 'KH', 'PA', 'CU', 'LA', 'CO', 'MT', 'UG',
                    'QA', 'PK', 'MK', 'DO']
    global var_1
    var_1 = tki.StringVar()
    var_1.set('Comedy')
    global var_2
    var_2 = tki.StringVar()
    var_2.set('до 60')
    global var_3
    var_3 = tki.StringVar()
    var_3.set('Арт-хаус')
    global var_4
    var_4 = tki.StringVar()
    var_4.set('Любая')
    global var_5
    var_5 = tki.StringVar()
    var_5.set('Любой')
    global tree
    style = ttk.Style(window_search_var1)
    style.theme_use("clam")
    style.configure("Treeview", background="gray21",
                    fieldbackground="gray21", foreground="#33CCFF")
    style.configure("Treeview.Heading", font=(None, 10))
    tree = ttk.Treeview(window_search_var1)
    tree['show'] = 'headings'
    tree["columns"] = ("c1", "c2", "c3")
    tree.column("c1", width=200)
    tree.column("c2", width=200)
    tree.column("c3", width=200)
    tree.heading("c1", text="Название фильмов")
    tree.heading("c2", text="Рейтинг Фильма")
    tree.heading("c3", text="Время")
    global sadtext
    sadtext = tki.Label(window_search_var1,
                        text="Мы не смогли найти вам желаемый фильм \n в нашей базе данных (个_个) ",
                        foreground="#33CCFF", background="gray23", font=("Gilroy ExtraBold", 15))

    global rep
    rep = True
    global printed_countries
    printed_countries = []

    """
    Эта функция нужна для проверки режима поиска фильмов:
    нужно ли выводить фильмы, которые были выведены ранее
    """

    def repeat():
        global rep
        rep = False

    def create_list_film():
        global printed_countries
        global rep
        if rep == False:
            printed_countries = []

        for i in tree.get_children():
            tree.delete(i)
        tree.pack()
        tree.pack_forget()
        sadtext.pack()
        sadtext.pack_forget()
        flag1 = 0
        # Здесь нужно чтобы

        pit = file
        data = pd.DataFrame(pit)
        data.sort_values(by='vote_average', ascending=False)
        index = 0
        jenre = var_1.get()
        time = var_2.get()
        budget = var_3.get()
        country = var_4.get()
        language = var_5.get()
        for i in range(len(data)):
            if data['budget'][i] < 500000:
                film_budget = 'Арт-хаус'
            else:
                if data['budget'][i] < 1000000:
                    film_budget = 'Низкобюджетное'
                else:
                    if data['budget'][i] < 30000000:
                        film_budget = 'Среднебюджетное'
                    else:
                        if data['budget'][i] < 200000000:
                            film_budget = 'Высокобюджетное'
                        else:
                            film_budget = 'Блокбастер'
            if data['runtime'][i] <= 60:
                runtime = 'до 60'
            else:
                if data['runtime'][i] <= 90:
                    runtime = 'от 60 до 90'
                else:
                    if data['runtime'][i] <= 120:
                        runtime = 'от 90 до 120'
                    else:
                        if data['runtime'][i] <= 150:
                            runtime = 'от 120 до 150'
                        else:
                            runtime = 'от 150'
            if rep == False:
                if (jenre in data['genres'][i] and
                        runtime in time and
                        film_budget in budget and
                        (country in data['production_countries'][i] or
                         country == 'Любая') and
                        (data['original_language'][i] == language or
                         language == 'Любой') and
                        index < 10):
                    flag1 = 1
                    tree.insert("", 0,
                                values=(data['original_title'][i], data['vote_average'][i], int(data['runtime'][i])))
                    index = index + 1
                    printed_countries.append(data['original_title'][i])

            else:
                if (data['original_title'][i] not in printed_countries and
                        jenre in data['genres'][i] and
                        runtime in time and
                        film_budget in budget and
                        (country in data['production_countries'][i] or
                         country == 'Любая') and
                        (data['original_language'][i] == language or
                         language == 'Любой') and
                        index < 10):
                    flag1 = 1
                    tree.insert("", 0,
                                values=(data['original_title'][i], data['vote_average'][i], int(data['runtime'][i])))
                    index = index + 1
                    printed_countries.append(data['original_title'][i])

        if flag1 == 0:
            sadtext.place(anchor="c", rely=0.8, relx=0.5)
        else:
            tree.place(relx=0.1, rely=0.5)
        rep = True

    main_menu_1 = tki.OptionMenu(window_search_var1, var_1, *jenre_list)
    main_menu_1.place(relx=0.04, rely=0.2)
    main_menu_1.config(bg="gray21", fg="#33CCFF", width=7, height=3, font=("Gilroy ExtraBold", 13))

    main_menu_2 = tki.OptionMenu(window_search_var1, var_2, *time_list)
    main_menu_2.place(relx=0.22, rely=0.2)
    main_menu_2.config(bg="gray21", fg="#33CCFF", width=9, height=3, font=("Gilroy ExtraBold", 13))

    main_menu_3 = tki.OptionMenu(window_search_var1, var_3, *budget_list)
    main_menu_3.place(relx=0.432, rely=0.2)
    main_menu_3.config(bg="gray21", fg="#33CCFF", width=14, height=3, font=("Gilroy ExtraBold", 13))

    main_menu_4 = tki.OptionMenu(window_search_var1, var_4, *country_list)
    main_menu_4.place(relx=0.703, rely=0.2)
    main_menu_4.config(bg="gray21", fg="#33CCFF", width=5, height=3, font=("Gilroy ExtraBold", 13))

    main_menu_5 = tki.OptionMenu(window_search_var1, var_5, *language_list)
    main_menu_5.place(relx=0.84, rely=0.2)
    main_menu_5.config(bg="gray21", fg="#33CCFF", width=5, height=3, font=("Gilroy ExtraBold", 13))

    style_forbutton = ttk.Style()
    style_forbutton.configure('W.TButton', font=('Gilroy ExtraBold', 15), foreground="#33CCFF", background="gray21")
    main_menu_6 = ttk.Button(window_search_var1, text="Найти фильмы", style='W.TButton',
                             command=create_list_film)
    main_menu_6.place(relx=0.5, rely=0.05, anchor="c")

    main_menu_7 = ttk.Button(window_search_var1, text="Сброс найденных фильмов", style='W.TButton',
                             command=repeat)
    main_menu_7.place(relx=0.5, rely=0.96, anchor="c")

    main_menu.add_command(label="    Жанр      ")
    main_menu.add_command(label="Длительность ")
    main_menu.add_command(label="Бюджет      ")
    main_menu.add_command(label="Страна    ")
    main_menu.add_command(label="Язык      ")


"""
Эта функция строит гистограмму, показывающую
зависимость между хронометражом фильма и его рейтингом
"""


def Duration_Rating():
    global file
    data = pd.DataFrame(file)
    x1 = list(data.runtime)
    y1 = list(data.vote_average)
    x2 = list()
    y2 = list()
    gr1 = dict(zip(x1, y1))
    for i in range(45464):
        s = 0
        j = 0
        k = 0
        for key in gr1:
            try:
                if int(key) == i:
                    s = s + gr1[key]
                    j = j + 1
                if j != 0:
                    k = s / j
            except:
                pass
        if i < 250 and j != 0 and k != 0:
            x2.append(i // 20 * 20)
            y2.append(k)
    sns.boxplot(x=x2, y=y2)
    plt.title("Зависимость рейтинга фильма от хронометража")
    plt.xlabel("Хронометраж")
    plt.ylabel("Рейтинг Фильма")
    plt.savefig("D:/Work/Graphics/Зависимость рейтинга фильма от хронометража.png", dpi=120)
    plt.show()


def Get_Relev():
    global file
    plt.hist(file['runtime'], color='blue', edgecolor='black', bins=int(701 / 10))
    plt.title("Отображение количества фильмов(ось y) \n в зависимости от хронометража(ось x) ", fontsize=10, loc='left')
    plt.savefig("D:/Work/Graphics/Релевантность, длительность фильмов, минут.png", dpi=105)
    plt.show()

    def funct1(ogr1, ogr2, atr):
        a1 = []
        for a in file[atr]:
            if (ogr1 < a < ogr2):
                a1.append(a)
        return (len(a1))

    nameofvaluesruns = [funct1(0, 52, 'runtime'), funct1(52, 200, 'runtime'), funct1(200, 706, 'runtime')]
    nameofindexruns = ["короткометр.", "полн.метр", "мини-сериалы"]
    plt.figure(figsize=(10, 5))
    plt.bar(nameofindexruns, nameofvaluesruns, color='maroon',
            width=0.4)
    plt.title("Отображение количества фильмов(ось y) \n в зависимости от хронометража категориально(ось x) ",
              fontsize=10, loc='left')
    plt.savefig("D:/Work/Graphics/Релевантность, метраж.png", dpi=70)
    plt.show()

    nameofvaluesbud = [funct1(0, 500000, 'budget'), funct1(500000, 1000000, 'budget'),
                       funct1(1000000, 30000000, 'budget'), funct1(30000000, 300000000, 'budget'),
                       funct1(300000000, 380000000, 'budget')]
    nameofindexbud = ["фестив.кино и артхаус", "малобюд.", "среднебюдж.", "высокобюдж.", "блокбастер"]
    plt.figure(figsize=(10, 5))
    plt.bar(nameofindexbud, nameofvaluesbud, color='maroon',
            width=0.4)
    plt.title("Отображение количества фильмов(ось y) \n в зависимости от бюджета категориально(ось x) ", fontsize=10,
              loc='left')
    plt.savefig("D:/Work/Graphics/Релевантность, бюджет.png", dpi=70)

    plt.show()
    goodinmoney = 0
    badinmoney = 0
    for i in file.itertuples():
        if (i[6] > i[2]):
            goodinmoney = goodinmoney + 1
        if (i[6] < i[2]):
            badinmoney = badinmoney + 1
    nameofindexpayback = ["окупив.", "не окупив."]
    plt.bar(nameofindexpayback, [goodinmoney, badinmoney], color='maroon',
            width=0.4)
    plt.title("Отображение количества фильмов(ось y) \n в зависимости от разницы бюджета и сборов(ось x) ", fontsize=10,
              loc='left')
    plt.savefig("D:/Work/Graphics/Релевантность, окупившиеся фильмы.png", dpi=100)

    plt.show()


def Get_Pivot_Table():
    data = pd.DataFrame(file)

    filtered_data = data[data["budget"] > 10000000]
    new_data = pd.pivot_table(filtered_data, index="original_title",
                              values=["budget", "revenue"], aggfunc=["mean"]).round(0)
    new_data.columns = new_data.columns.map("_".join)
    new_data.to_excel("D:/Work/Output/Высокобюджетные фильмы и их сборы.xlsx")


######################################################################################

"""Эта функция позволяет получить соотношение жанров фильмов в разных странах
мира, она примнимает на вход датасет и название страны(либо "All world" если
нужно узнать соотноешение в целом по датасету)"""


def Genres_Of_Cinema_Russia(Name_Of_Country="Russia"):
    global file
    Dct_For_Genres = {}
    example_genres = ['Comedy', 'Animation', 'Horror', 'Fantasy', 'Thriller', 'Action', 'Romance', 'Mystery',
                      'Adventure', 'Drama', 'Western', 'Crime', 'Family']
    namesofgenres = []
    for index, row in file.iterrows():
        if Name_Of_Country == 'All world':
            for a in example_genres:
                if a in row['genres']:
                    if a in Dct_For_Genres:
                        Dct_For_Genres[a] += 1
                    else:
                        Dct_For_Genres[a] = 1
                        namesofgenres.append(a)
        else:
            if Name_Of_Country in row['production_countries']:
                for a in example_genres:
                    if a in row['genres']:
                        if a in Dct_For_Genres:
                            Dct_For_Genres[a] += 1
                        else:
                            Dct_For_Genres[a] = 1
                            namesofgenres.append(a)

    figforgenres = plt.figure(figsize=(20, 7))
    valuesofgenres = []
    if Name_Of_Country == 'All world':
        figforgenres.legend(title='Genres of films all over the world', loc="upper center")
    else:
        figforgenres.legend(title='Соотношение жанров фильмов, снятых в России', loc="upper center")
    for a in namesofgenres:
        valuesofgenres.append(Dct_For_Genres[a])
    plt.pie(valuesofgenres, labels=namesofgenres)
    plt.savefig("D:/Work/Graphics/Жанры фильмов в России.png")
    plt.show()


def Genres_Of_Cinema_China(Name_Of_Country="China"):
    global file
    Dct_For_Genres = {}
    example_genres = ['Comedy', 'Animation', 'Horror', 'Fantasy', 'Thriller', 'Action', 'Romance', 'Mystery',
                      'Adventure', 'Drama', 'Western', 'Crime', 'Family']
    namesofgenres = []
    for index, row in file.iterrows():
        if Name_Of_Country == 'All world':
            for a in example_genres:
                if a in row['genres']:
                    if a in Dct_For_Genres:
                        Dct_For_Genres[a] += 1
                    else:
                        Dct_For_Genres[a] = 1
                        namesofgenres.append(a)
        else:
            if Name_Of_Country in row['production_countries']:
                for a in example_genres:
                    if a in row['genres']:
                        if a in Dct_For_Genres:
                            Dct_For_Genres[a] += 1
                        else:
                            Dct_For_Genres[a] = 1
                            namesofgenres.append(a)

    figforgenres = plt.figure(figsize=(20, 7))
    valuesofgenres = []
    if Name_Of_Country == 'All world':
        figforgenres.legend(title='Genres of films all over the world', loc="upper center")
    else:
        figforgenres.legend(title='Соотношение жанров фильмов, снятых в Китае', loc="upper center")
    for a in namesofgenres:
        valuesofgenres.append(Dct_For_Genres[a])
    plt.pie(valuesofgenres, labels=namesofgenres)
    plt.savefig("D:/Work/Graphics/Жанры фильмов в Китае.png")
    plt.show()


def Genres_Of_Cinema_USA(Name_Of_Country="United States of America"):
    global file
    Dct_For_Genres = {}
    example_genres = ['Comedy', 'Animation', 'Horror', 'Fantasy', 'Thriller', 'Action', 'Romance', 'Mystery',
                      'Adventure', 'Drama', 'Western', 'Crime', 'Family']
    namesofgenres = []
    for index, row in file.iterrows():
        if Name_Of_Country == 'All world':
            for a in example_genres:
                if a in row['genres']:
                    if a in Dct_For_Genres:
                        Dct_For_Genres[a] += 1
                    else:
                        Dct_For_Genres[a] = 1
                        namesofgenres.append(a)
        else:
            if Name_Of_Country in row['production_countries']:
                for a in example_genres:
                    if a in row['genres']:
                        if a in Dct_For_Genres:
                            Dct_For_Genres[a] += 1
                        else:
                            Dct_For_Genres[a] = 1
                            namesofgenres.append(a)

    figforgenres = plt.figure(figsize=(20, 7))
    valuesofgenres = []
    if Name_Of_Country == 'All world':
        figforgenres.legend(title='Genres of films all over the world', loc="upper center")
    else:
        figforgenres.legend(title='Соотношение жанров фильмов, снятых в США', loc="upper center")
    for a in namesofgenres:
        valuesofgenres.append(Dct_For_Genres[a])
    plt.pie(valuesofgenres, labels=namesofgenres)
    plt.savefig("D:/Work/Graphics/Жанры фильмов в США.png")
    plt.show()


def Genres_Of_Cinema(Name_Of_Country="All world"):
    global file
    Dct_For_Genres = {}
    example_genres = ['Comedy', 'Animation', 'Horror', 'Fantasy', 'Thriller', 'Action', 'Romance', 'Mystery',
                      'Adventure', 'Drama', 'Western', 'Crime', 'Family']
    namesofgenres = []
    for index, row in file.iterrows():
        if Name_Of_Country == 'All world':
            for a in example_genres:
                if a in row['genres']:
                    if a in Dct_For_Genres:
                        Dct_For_Genres[a] += 1
                    else:
                        Dct_For_Genres[a] = 1
                        namesofgenres.append(a)
        else:
            if Name_Of_Country in row['production_countries']:
                for a in example_genres:
                    if a in row['genres']:
                        if a in Dct_For_Genres:
                            Dct_For_Genres[a] += 1
                        else:
                            Dct_For_Genres[a] = 1
                            namesofgenres.append(a)

    figforgenres = plt.figure(figsize=(20, 7))
    valuesofgenres = []
    if Name_Of_Country == 'All world':
        figforgenres.legend(title='Соотношение жанров фильмов, снятых по всему миру', loc="upper center")
    else:
        figforgenres.legend(title='Genres of films in ' + Name_Of_Country, loc="upper center")
    for a in namesofgenres:
        valuesofgenres.append(Dct_For_Genres[a])
    plt.pie(valuesofgenres, labels=namesofgenres)
    plt.savefig("D:/Work/Graphics/Жанры фильмов по всему миру.png")
    plt.show()


#                             ДО СЮДОВА - ВСТАВЛЯТЬ БЕЗДУМНО

################################################################################

def dataset_popen():
    os.startfile("dc/Work/Data/moviedataset.xlsx")


def user_guide_popen():
    os.startfile("D:/Work/Notes/Руководство пользователя.txt")


def dev_guide_popen():
    os.startfile("D:/Work/Notes/Руководство разработчика.txt")


#                             ОТСЮДОВА
def show_relev_reports():
    graphics_window = tki.Toplevel(root)
    graphics_window.geometry("750x650")
    graphics_window.resizable(width=0, height=0)
    graphics_window.wm_attributes("-topmost", 1)
    graphics_label = tki.Label(graphics_window, text="Релевантность сборки: графики")
    graphics_label.pack()

    Get_Relev()

    pimage1 = tki.PhotoImage(file="D:/Work/Graphics/Релевантность, бюджет.png")
    pimage2 = tki.PhotoImage(file="D:/Work/Graphics/Релевантность, длительность фильмов, минут.png")
    pimage3 = tki.PhotoImage(file="D:/Work/Graphics/Релевантность, метраж.png")
    pimage4 = tki.PhotoImage(file="D:/Work/Graphics/Релевантность, окупившиеся фильмы.png")

    canvas = tki.Canvas(graphics_window, bg="white", width=650, height=450)
    image_list = [pimage1, pimage2, pimage3, pimage4]
    global count, i1
    count = 0
    canvas.image = image_list[count]
    canvas.pack()
    i1 = canvas.create_image(360, 260, anchor="center", image=image_list[count])

    def next_image():
        global count, i1
        canvas.delete(i1)
        count += 1
        prev_button.config(state="active")
        canvas.image = image_list[count]
        i1 = canvas.create_image(360, 260, anchor="center", image=image_list[count])

        if count == 3:
            next_button.config(state="disabled")

    def prev_image():
        global count, i1
        canvas.delete(i1)
        count -= 1
        next_button.config(state="active")
        canvas.image = image_list[count]
        i1 = canvas.create_image(360, 260, anchor="center", image=image_list[count])

        if count == 0:
            prev_button.config(state="disabled")

    next_button = tki.Button(graphics_window, text="Следующий", width=12, height=2, command=next_image)
    next_button.pack()
    prev_button = tki.Button(graphics_window, text="Предыдущий", width=12, height=2, command=prev_image)
    prev_button.pack()
    exit_button = tki.Button(graphics_window, text="Выход", width=12, height=2, command=graphics_window.destroy)
    exit_button.pack()


def show_graphic_reports():
    graphics_window = tki.Toplevel(root)
    graphics_window.geometry("750x650")
    graphics_window.resizable(width=0, height=0)
    graphics_window.wm_attributes("-topmost", 1)
    graphics_label = tki.Label(graphics_window, text="Графические отчёты")
    graphics_label.pack()

    answer = messagebox.askyesno(title="Подтверждение", message="Продолжить? (придётся немного подождать)")
    if not answer:
        graphics_window.destroy()
        return
    Duration_Rating()
    Genres_Of_Cinema_China()
    Genres_Of_Cinema_Russia()
    Genres_Of_Cinema_USA()
    Genres_Of_Cinema()

    pimage1 = tki.PhotoImage(file="D:/Work/Graphics/Зависимость рейтинга фильма от хронометража.png")
    pimage2 = tki.PhotoImage(file="D:/Work/Graphics/Жанры фильмов по всему миру.png")
    pimage3 = tki.PhotoImage(file="D:/Work/Graphics/Жанры фильмов в России.png")
    pimage4 = tki.PhotoImage(file="D:/Work/Graphics/Жанры фильмов в Китае.png")
    pimage5 = tki.PhotoImage(file="D:/Work/Graphics/Жанры фильмов в США.png")

    canvas = tki.Canvas(graphics_window, bg="white", width=1000, height=500)
    image_list = [pimage1, pimage2, pimage3, pimage4, pimage5]
    global count, i1
    count = 0
    canvas.image = image_list[count]
    canvas.pack()
    i1 = canvas.create_image(370, 260, anchor="center", image=image_list[count])

    def next_image():
        global count, i1
        canvas.delete(i1)
        count += 1
        prev_button.config(state="active")
        canvas.image = image_list[count]
        i1 = canvas.create_image(370, 260, anchor="center", image=image_list[count])

        if count == 4:
            next_button.config(state="disabled")

    def prev_image():
        global count, i1
        canvas.delete(i1)
        count -= 1
        next_button.config(state="active")
        canvas.image = image_list[count]
        i1 = canvas.create_image(370, 260, anchor="center", image=image_list[count])

        if count == 0:
            prev_button.config(state="disabled")

    next_button = tki.Button(graphics_window, text="Следующий", width=12, height=2, command=next_image)
    next_button.pack()
    prev_button = tki.Button(graphics_window, text="Предыдущий", width=12, height=2, command=prev_image)
    prev_button.pack()
    prev_button.config(state="disabled")
    exit_button = tki.Button(graphics_window, text="Выход", width=12, height=2, command=graphics_window.destroy)
    exit_button.pack()


def show_text_reports():
    Get_Pivot_Table()
    os.startfile("D:/Work/Output/Высокобюджетные фильмы и их сборы.xlsx")


#                                                         ДО СЮДОВА

#########################################################################################################################

root = tki.Tk()
root.wm_attributes("-topmost", 1)  # Запуск поверх всех окон
root.title("CinemaEvening")
root.geometry('700x500+50+50')
root.resizable(width=0, height=0)
root["bg"] = "gray23"
logotext = tki.Label(root, text="Киновечер", foreground="#33CCFF", background="gray23", font=("Gilroy ExtraBold", 30))
logotext.place(relx=0.35, rely=0.4)
under_logotext = tki.Label(root, text="Когда нужно подобрать фильм и не только...", foreground="white",
                           background="gray23", font=("Gilroy ExtraBold", 10))
under_logotext.place(relx=0.305, rely=0.5)
mainmenu = tki.Menu(root, tearoff=0, foreground="#33CCFF", background="gray21", relief="ridge", borderwidth=2)

mainmenu_1 = tki.Menu(mainmenu, tearoff=0, foreground="#33CCFF", background="gray21", relief="ridge", borderwidth=2)
mainmenu_1.add_command(label="Подбор фильма по категориям", command=search_film)

mainmenu_2 = tki.Menu(mainmenu, tearoff=0, foreground="#33CCFF", background="gray21", relief="ridge", borderwidth=2)
mainmenu_2.add_command(label="Графические отчеты", command=show_graphic_reports)
mainmenu_2.add_command(label="Текстовые отчеты", command=show_text_reports)
mainmenu_2.add_command(label="Отображение релевантности выборки", command=show_relev_reports)

mainmenu_3 = tki.Menu(mainmenu, tearoff=0, foreground="#33CCFF", background="gray21", relief="ridge", borderwidth=2)
mainmenu_3.add_command(label="Руководство пользователя", command=user_guide_popen)
mainmenu_3.add_command(label="Руководство разработчика", command=dev_guide_popen)
mainmenu_3.add_command(label="База данных", command=dataset_popen)

mainmenu.add_cascade(label="О проекте", menu=mainmenu_3, background="gray21")
mainmenu.add_cascade(label="Аналитика и разносторонний анализ", menu=mainmenu_2)
mainmenu.add_cascade(label="Что сегодня посмотреть?", menu=mainmenu_1)
root.config(menu=mainmenu)

root.mainloop()