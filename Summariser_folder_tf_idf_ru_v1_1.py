'''
Определение ключевых слов по tfidf для документов
из набора русскоязычных текстов, помещенных в папку Samples,
с последующим построением Абстракта для каждого документа,
на основании этих уникальных ключевых слов.
В перспективе нужно сделать вывод их в таблицу Эксель построчно
в формате "Название файла : Абстракт"

В этой версии в качестве слов для подсчета tfidf выбраны только сущ. и глаг.
'''
import os
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import math
import gensim
from gensim.utils import simple_preprocess
from collections import Counter
import xlsxwriter
import textract
import pymorphy2
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

morph = pymorphy2.MorphAnalyzer()

def compute_tfidf(corpus):
    '''на входе - список со списками лемматизированных слов для каждого текста,
    на выходе - словарь "слово:tfidf" для каждого слова'''

    def compute_tf(text):
        '''На входе - список лемматизированных слов отдельного текста,
        на выходе - словарь "слово: частота" для этого текста'''
        tf_text = Counter(text)# Делает словарь "элемент списка:кол-во употреблений" 
        for i in tf_text:
            tf_text[i] = tf_text[i]/float(len(text))
        return tf_text

    def compute_idf(word, corpus):
        '''на входе - конкретное слово и список со списками слов каждого текста,
        на выходе - idf этого слова'''
        return math.log10(len(corpus)/sum([1.0 for i in corpus if word in i]))

    documents_list = []
    for text in corpus:
        tf_idf_dictionary = {}
        computed_tf = compute_tf(text)
        for word in computed_tf:
            tf_idf_dictionary[word] = computed_tf[word] * compute_idf(word, corpus)
        documents_list.append(tf_idf_dictionary)
    return documents_list

def lemmat_ru(file):
    '''Принимает текст, разбивает на слова, нормализирует каждое слово.
    На выходе - список нормализированных слов
    Здесь нет выделения только существительных и глаголов!!!
    '''
    list1=[]
    words1=simple_preprocess(file)
    for word in words1:
        p=morph.parse(word)[0]
        if 'NOUN' in p.tag or 'VERB' in p.tag:
            list1.append(p.normal_form)
    return list1

def value_sentence(sentences):
    '''Ф-ция вычисляет ценность каждого предложения
    в зависимости от количества ключевых слов в нем, и выдает словарь "предложение:ценность"'''
    sentenceValue = dict() # Создаем пустой словарь для пар «предложение : его весомость»
    for sentence in sentences:
        for word, freq in freqTable.items():
            if word in sentence.lower(): # Проверяем каждое ключ. слово на наличие в этом предложении
                if sentence in sentenceValue:
                    sentenceValue[sentence] += freq # весомость предл. увеличивается на част. слова
                    #sentenceValue[sentence]= sentenceValue[sentence]/len(sentence.split(' '))
                else:
                    sentenceValue[sentence] = freq # весомость предложения равна частоте слова
                    #sentenceValue[sentence]= sentenceValue[sentence]/len(sentence.split(' '))
    return sentenceValue

def average_value(sentenceValue):# Average value of a sentence from the original text
    sumValues = 0
    for sentence in sentenceValue:  #Для каждого предложения из словаря «предложение : его весомость»
        sumValues += sentenceValue[sentence] # Вычисляем суммарную стоимость всех предложений
    try:
        average = sumValues / len(sentenceValue) # Суммарную стоимость делим на количество предложений
    except ZeroDivisionError:
        average=1
    return average


abs_filenames=[] #список абсолютных имен файлов с полным путем
corpus=[]
filenames=[]
results=[]

file_path = filedialog.askdirectory()
''' Открываем каждый из файлов в папке, читаем текст,
разбиваем текст на слова и лемматизируем их функцией lemmat_ru,
добавляем список нормализованных слов текста в общий список списков corpus,
а имена файлов добавляем в общий список имен filenames'''
for filename in os.listdir(file_path):
    abs_filenames.append(os.path.normpath(os.path.join(file_path, filename))) 
    filenames.append(filename)
    extension = os.path.splitext(filename)[1][1:] # Extract the extension from the filename
    if extension == "txt":
        with open(os.path.join(file_path, filename), 'r', encoding='utf-8') as f:
            text = f.read()
            words=lemmat_ru(text)
            corpus.append(words)
    elif extension == "docx": # use the textract library to open it. 
        text = textract.process(os.path.join(file_path, filename))
        text = text.decode("utf8")
        words=lemmat_ru(text)
        corpus.append(words)
    else:
        continue
'''
Две следующие вставки - на случай, если результат нужно выводить в Эксель
# Открываем Экселевский файл, добавляем лист
workbook = xlsxwriter.Workbook('exp1.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write_column('A1', abs_filenames)
worksheet.write_column('B1', filenames) # Названия файлов - в первую колонку
'''
x=compute_tfidf(corpus)
for element in x:
    results.append(sorted(element, key=element.get)[-8:]) #общий список со списками ключевых слов

'''
# Каждый список ключ.слов превращаем в тип "строка", чтобы записывать в Эксель
i=0 #номер строки в таблице
for element in results:
    x=', '.join(element)
    worksheet.write(i,2, x) # в третью колонку, по одной строке
    i=i+1
workbook.close()
'''

'''Теперь у нас есть список abs_filenames с именами файлов
и список results с ключевыми словами для них.
Открывая по очереди каждый файл из списка файлов мы можем делать саммари,
опираясь на ключевые слова из соответствующего элемента списка ключевых слов'''


for i in range(0, len(results)):
    filename=abs_filenames[i]
    relev_words=results[i]

    with open(filename, 'r', encoding='utf-8') as f:
        text = f.read()
        listW=lemmat_ru(text)
        listRelWords=[]
        for word in listW:
            if word in relev_words:
                listRelWords.append(word)
        freqTable=Counter(listRelWords)
        print(freqTable)
# слово "быть" появляется во всех списках. Проверить алгоритм tf_idf !!!
'''
        sentences = sent_tokenize(text) # Токенизируем по предложениям, получим список
        n=len(sentences)

        if n>8: # саммари иммет смысл, если текст больше 8 предложений
        
            sentenceValue=value_sentence(sentences) # ценность каждого предложения            
            averageValue=average_value(sentenceValue)# средний вес предложения в тексте
              
            k=0.8 # коэфф-нт важности, относительно среднего веса предложения с keywords
            
            while n>8: # Ограничим Саммари восемью предложениями, повышая коэфф. важности
                k=k+0.05
                summary = ''
                for sentence in sentences:
                    if (sentence in sentenceValue) and (sentenceValue[sentence] > (k * averageValue)):
                        summary += "\n" + sentence # Добавляем это ценное предложение в конец списка 
                sentSum=sent_tokenize(summary)
                n=len(sentSum)
            
        else:
            summary =text
        print('\n', filename, '\n', summary)
'''
