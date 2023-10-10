from bs4 import BeautifulSoup
import urllib
import requests
import re
import pandas as pd

import openpyxl as op
import pandas as pd

import pyphen
from collections import Counter


def text_extractor(url, title):
    res=requests.get(url)
    h=res.content
    soup=BeautifulSoup(h, 'html.parser')
    
    sx=''
    t=''
    slt=soup.find_all('p')
    
    for a in slt:
        if a.has_attr('title') or a.has_attr('id') or a.has_attr('class') or a.has_attr('href') or a.has_attr('strong'):
            continue
        else:
            sx+=(str(a))+'\t'

    t=BeautifulSoup(sx, features='html.parser').text

    g=open(title, 'w',encoding='utf-8')
    g.write(t)
    g.close()



with open('Docs/positive-words.txt','r') as file:
    pos=file.read().split()

with open('Docs/negative-words.txt','r') as file:
    neg=file.read().split()


def listify(filename):
    try:
        
        with open(filename, 'r', encoding='utf-8') as file:
            words = file.read().split()
        if len(words)==0:
            with open(filename, 'r') as file:
                words = file.read().split()
        return set(words)
        
    except UnicodeEncodeError:
        print('FacedError')
        with open(filename, 'r') as file:
            words = file.read().split()
        if len(words)==0:
            with open(filename, 'r') as file:
                words = file.read().split()
        return set(words)
        

def remove_common_words(inp, st_words):
    A = listify(inp)

    pp_c=0
    pp = ['i', 'my', 'we','ours','us','I','My','We','Ours','Us']
    for i in pp:
        if i in A and i!='US':
            pp_c+=1

    
    B = listify(st_words)
    filtered = A-B

    word_c=len(A)-len(filtered)
    try:
            
        with open('Filtered/'+inp, 'w') as file:
            file.write('\n'.join(filtered))
    except UnicodeEncodeError:
        with open('Filtered/'+inp, 'w', encoding='utf-8') as file:
            file.write('\n'.join(filtered))

    ###############################################################
    #Calculating Positive, Negative, words
    try:
        with open('Filtered/'+inp, 'r') as file:
            g=file.read().split()
    except UnicodeDecodeError:
        with open('Filtered/'+inp, 'r', encoding='utf-8') as file:
            g=file.read().split()
    no_pos, no_neg=0,0
    for i in g:
        if i in pos:
            no_pos+=1
        if i in neg:
            no_neg+=1
    polarity_score=(no_pos-no_neg)/((no_pos+no_neg)+0.000001)
    subjectivity_score=(no_pos+no_neg)/(len(filtered)+0.000001)

    return word_c, no_pos, no_neg, polarity_score,subjectivity_score, pp_c
    
def calculate_metrics(file_path):
    try:
        with open(file_path, 'r') as infile:
            text = infile.read()
    except:
        with open(file_path, 'r', encoding='utf-8') as infile:
            text = infile.read()

    
    if len(text)==0:
        return 0,0,0,0,0,0,0,0,0
    else:
            
        syllables_per_word=0
        for i in "aeiouAEIOU":
            syllables_per_word+=text.count(i)
        syllables_per_word-=text.count('ed')+text.count('es')
    
        
        sentences = re.split(r'[.!?]', text)
        sentences = [sentence.strip() for sentence in sentences if sentence.strip()]
    
        dic = pyphen.Pyphen(lang='en')
    
        total_sentences = len(sentences)
        total_words = 0
        total_complex_words = 0
        total_word_length = 0
        total_personal_pronouns = 0
        total_syllables = 0
    
        for sentence in sentences:
            words = sentence.split()
            total_words += len(words)
            total_word_length += sum(len(word) for word in words)
    
            for word in words:
                syllable_count = len(dic.inserted(word).split('-'))
                total_syllables += syllable_count
    
                if syllable_count > 2:
                    total_complex_words += 1
    
        #print(total_words)
        
        average_sentence_length = total_words / total_sentences if total_sentences > 0 else 1
        complex_word_percentage = (total_complex_words)
        average_word_length = total_word_length / total_words if total_words > 0 else 1
        fog_index=0.4*(average_sentence_length+complex_word_percentage)
        
        average_syllables_per_word = total_syllables / total_words if total_words > 0 else 1
        avg_words_per_sentence=total_words/len(sentences) if len(sentences)>0 else 1
    
        return average_sentence_length, complex_word_percentage,average_word_length,personal_pronoun,average_syllables_per_word,avg_words_per_sentence,syllables_per_word/total_words,fog_index, total_complex_words

def uploader(x,l,file_path):
 
    workbook = op.load_workbook(file_path)
    sheet = workbook.active
    for i in range(13):
        a=sheet.cell(row=x, column=i+3)
        a.value=l[i]
        
    workbook.save('output copy.xlsx')


f=pd.read_excel("Docs/Input.xlsx")

url=list(f.URL)
url_id=list(f.URL_ID)

for i in range(len(url)):
    text_extractor(url[i], str(url_id[i])+'.txt')

stop_words = 'Docs/StopWords.txt'
x=2
for i in url_id:
    
    if i!=44:
        input_file = str(i)+'.txt'
        word_count,positive_words, negative_words, polarity_score, subjectivity_score, personal_pronoun = remove_common_words(input_file, stop_words)
        inp = 'Filtered/'+input_file
        
        avg_sentence_length, complex_word_percentage, avg_word_length, personal_pronoun, avg_syllables_per_word, avg_words_per_sentence, syllables_per_word, fog_index, total_complex_words= calculate_metrics(inp)
        l=[positive_words,negative_words,polarity_score,subjectivity_score,avg_sentence_length,complex_word_percentage,fog_index,avg_words_per_sentence,total_complex_words, word_count,syllables_per_word, personal_pronoun, avg_word_length]
        #uploading the derived values into the csv file
        uploader(x, l, "Docs/Output Data Structure.xlsx")
        x+=1
