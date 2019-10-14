# coding: utf-8

import os
import regex
import xlsxwriter

stopwords = ['на','не','для','по','за','из-за','над','про','через','после','вокруг','около','без','от','до','возле', \
             'перед','об','при','позади','вверх','вверху','сверху','вниз','внизу','снизу','затем','слева', \
             'справа','левее','правее','налево','направо','ввиду','наподобие','вроде','вместо','насчет','насчёт', \
             'вследствие','вслед','внутри','снаружи','навстречу','вдоль','несмотря','невзирая','из-под','то', \
             'ого','ага','нет','да','бы','так','это','но','также','почти','возможно','возможен','возможна','возможны', \
             'он','ему','его','него','им','нем','нём', \
             'она','ее','её','ней','нее','неё','ей', \
             'они','их','им','них','ими','мы','нас','нам','нами', \
             'ты','тебя','тебе','тобой','вы','вас','вам','вами', \
             'меня','мне','меня','мной','все','всех','всем','всеми', \
             'сколько','кто','как','где','когда','что','чего','кого','кому','чему','кем','чем','чём','почему','зачем', \
             'потому','затем','чтобы','того','этого','тот','этот','тому','этому','тем','этим','если','который', \
             'которого','которому','которым','котором','которая','которой','которую','которые','которых','которым', \
             'которыми','эта','этой','эту','теми','тех','себе','себя','собой','некоторый','некоторого','некоторому',\
             'некоторым','некотором','каждый','каждого','каждому','каждым','каждом','каждая','каждой','каждую', \
             'каждое','каждые','каждых','каждым','каждыми', \
             'один','два','три','четыре','пять','шесть','семь','восемь','девять', \
             'десять','двадцать','тридцать','сорок','пятьдесят','шестьдесят','семьдесят','восемьдесят','девяносто', \
             'сто','двести','триста','четыреста','пятьсот','шестьсот','семьсот','восемьсот','девятьсот', \
             'тысяча','миллион','миллиард' \
             'первый', 'второй', 'третий', 'четвертый', 'пятый', 'шестой', 'седьмой', 'восьмой', 'девятый', \
             'десятый','двадцатый','тридцатый','сороковой','пятидесятый','шестидесятый','семидесятый','восмидесятый', \
             'девяностый','сотый','двухсотый','трехсотый','трёхсотый','четырехсотый','четырёхсотый','пятисотый', \
             'шестисотый', 'семисотый', 'восьмисотый', 'девятисотый', \
             'тысячный', 'миллионный', 'миллиардный', \
             'можно','возможно','никакой','какой','какого','какому','какого', \
             'каким','каком','какая','какую','какие','каких','какими']

stopwordsdictpath = input("Введите путь к словарю стоп-слов (список слов, разделенных запятыми): ")
if not os.path.exists(stopwordsdictpath):
    print("Указанный файл не существует! Использую встроенный словарь")
else:
    stopwordsdictfile = open(stopwordsdictpath, 'r')
    stopwordsdict = stopwordsdictfile.read()
    stopwordsdictfile.close()
    stopwordsdict = regex.sub('\s', ",", stopwordsdict)  # заменяю пробельные символы запятыми
    print(stopwordsdict)
    stopwordsdict = regex.sub('[^А-Яа-яA-Za-z,]', "", stopwordsdict)  # удаляю все символы, кроме букв и запятых
    print(stopwordsdict)
    stopwords = stopwordsdict.split(',')
    print(stopwords)

# выбор файла с текстом

fullname = input("Введите путь к файлу: ")
if not os.path.exists(fullname):
    print("Указанный файл не существует!")
else:
    dirname = os.path.dirname(fullname)
    filename = os.path.basename(fullname)
    thefile = open(fullname,"r")
    text = thefile.read()
    thefile.close()

# Удаляю лишнее из текста
    text = regex.sub('\s'," ",text)  # заменяю пробельные символы простым пробелом
    text = regex.sub('[^А-Яа-яA-Za-z -]', "", text)  # удаляю все символы, кроме букв и пробелов
    text = regex.sub(' о ком ',' ',text)
    text = regex.sub(' во-[А-Яа-я]+ ', ' ', text)
    text = regex.sub('-то ', ' ', text)
    text = regex.sub(' [А-Яа-я]+-нибудь ', ' ', text)
    text = regex.sub('-ка ', ' ', text)
    text = regex.sub('-таки ', ' ', text)
    raw_words = text.split(' ')

    words = []
    for word in raw_words:
       if len(word) > 1 and not word.lower() in stopwords:
           words.append(word.lower())

# Ищу ключевые слова
    wordcount = {}
    for word in words:
        if wordcount.get(word) == None:
            r = {word : words.count(word)}
            wordcount.update(r)

    resultlist = []
    for i in wordcount.items():
        resultlist.append(i)
    sortedresult = sorted(resultlist, key=lambda x: x[1], reverse=True)

# Записываю результат в файл
    xlsxfullname = fullname + '.xlsx'
    overwrite = 'нет'
    if os.path.exists(xlsxfullname):
        print("Файл ",xlsxfullname, " уже существует!")
        overwrite = input('Введите "Да", чтобы перезаписать его: ')
        if  overwrite == 'Да':
            try:
                os.remove(xlsxfullname)
            except PermissionError:
                print("Закройте файл и повторите попытку!")
                overwrite = 'нет'
    else:
        overwrite = 'Да'

    if overwrite == 'Да':
        workbook = xlsxwriter.Workbook(xlsxfullname)
        worksheet = workbook.add_worksheet()

        row = 0
        col = 0
        worksheet.write(row, col, 'Слово')
        worksheet.write(row, col + 1, 'Частота употребления')
        row += 1
        for word,count in sortedresult:
            worksheet.write(row, col, word)
            worksheet.write(row, col + 1, count)
            row += 1

        workbook.close()