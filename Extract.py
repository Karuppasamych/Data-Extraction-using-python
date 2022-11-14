import nltk
nltk.download('stopwords')
nltk.download('punkt')
from rake_nltk import Rake
from nltk import tokenize
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
stop_words = set(stopwords.words('english'))
import operator

from operator import itemgetter
import xlsxwriter
import math

from collections import Counter




import pandas as pd


r = Rake()

# excelfile= "C:\\Users\\Techno\\Desktop\\Data_file\\OriginalData\\cleaned_notes.xlsx" #ARyabhata Auto Populate Report #data1 #C:\Users\Techno\Desktop\Data_file\OriginalData
excelfile= "C:\\Users\\101053\\Desktop\\Techno\\Data_file\\modernizing medicine billing services, LLC\\Modernizing Medicine Billing Services,LLC location id 186 practice id 412.xlsx" #ARyabhata Auto Populate Report #data1 #C:\Users\Techno\Desktop\Data_file\OriginalData

df = pd.read_excel(excelfile, sheet_name=0, usecols=['AgentNotes'], engine="openpyxl") #Testclean


keywords = []
cleaned_notes = []
importantkeywords = []

for index, row in df.items():

    iterdata =  df['AgentNotes'].tolist()

    for idata in iterdata:
        
        data = str(idata)
        
        text_tokens = word_tokenize(data)
        
        tokens_without_sw = [word for word in text_tokens if not word in stopwords.words()] #using stopkeywords
        
        filtered_sentence = (" ").join(tokens_without_sw)
        
        cleaned_notes.append(filtered_sentence)
        
        r.extract_keywords_from_text(data)

        rank = r.get_ranked_phrases_with_scores() # getting rank 
        
        test = []
        
        for keyword in rank:

            keyword_updated  = keyword[1].split()


            keyword_updated_string = " ".join(keyword_updated[:3])
            
            
            test.append(keyword_updated_string)
        
        keywords.append(test)
        
        
        # Writing the clean notes in excelsheet 
        workbook = xlsxwriter.Workbook('C:\\Users\\101053\\Desktop\\Techno\\Data_file\\ResultData\\billing_services2.xlsx')  # create/read the file
        
        worksheet = workbook.add_worksheet() # adding a worksheet
        
        row1 = 0
        col = 0
        
        for row in keywords:
        
            worksheet.write(col,row1, repr(row))# write repr(lst) to cell A1
        
            col +=1
        
        workbook.close()   
    
    
    # Finding length of each keyword and append to importantkeywords array
    for keys in keywords:
        
        for importkey in keys:
        
            if len(importkey) >5:
        
                importantkeywords.append(importkey)

        

selectivekeywords = []


# Extracting the top keywords

def topkeywords(a):    

    words = {}

    for data in a:

        if data in words:

            words[data] +=1

        else:

            words[data] =1

    # print(words,'WORDS=======')

    # Descending the topkeywords by count
    topwords = sorted(words.items(),key=operator.itemgetter(1),reverse=True)


    # Writing the top 30 keywords in excelsheet 
    
    df = pd.DataFrame(topwords)

    df.head(30).to_excel('C:\\Users\\101053\\Desktop\\Techno\\Data_file\\ResultData\\30topkeywords.xlsx')


    # Storing the count of important keywords
    data = {'eob received':0,'claim billed':0, 'adjust cpt':0, 'claim rebilled':0, 'claim submitted':0, 'therefore adjusted':0,'claim denied':0, 'therefore assigned':0, 'eob found':0, 'claim submitted':0, 'requested eob':0 , 'therefore claim':0, 'per previous notes':0}
    topwords = {}
    
    for word, numitem in words.items():

        # print(numitem,'NUMBER=======')

        if len(word) > 6:

            if numitem > 250:
                selectivekeywords.append(word)

        
        val = words.get(word)

        if 'eob received' in word:
        
            data['eob received'] += val
        
        if 'claim billed' in word:
        
            data['claim billed'] += val
        
        if 'adjust cpt' in word:
        
            data['adjust cpt'] += val

        if 'provide manual' in word:
        
            data['provide manual'] += val
        
        if 'claim rebilled' in word:
        
            data['claim rebilled'] += val
        
        if "claim submitted" in word:
        
            data['claim submitted'] +=val
        
        if "therefore adjusted" in word:
        
            data['therefore adjusted'] +=val

        if 'claim denied' in word:
        
            data['claim denied'] += val
        
        if 'therefore assigned' in word:
        
            data['therefore assigned'] += val
        
        if 'eob found' in word:
        
            data['eob found'] += val

        if 'requested eob' in word:
        
            data['requested eob'] += val
        
        if 'per previous notes' in word:
        
            data['per previous notes'] += val
        
        if "claim submitted" in word:
        
            data['claim submitted'] +=val
        
        if "therefore claim" in word:
        
            data['therefore claim'] +=val


    workbook = xlsxwriter.Workbook('C:\\Users\\101053\\Desktop\\Techno\\Data_file\\ResultData\\topkeywords.xlsx')  # create/read the file
    
    worksheet = workbook.add_worksheet() # adding a worksheet
    
    row1 = 0
    col = 0
    
    for row in selectivekeywords:
    
        worksheet.write(col,row1, repr(row))# write repr(lst) to cell A1
    
        col +=1
    
    workbook.close()
    
    print(data, 'COUNT')

topkeywords(importantkeywords)

# print(selectivekeywords,'SELECTVE')

# print(len(selectivekeywords),'LEN====')



