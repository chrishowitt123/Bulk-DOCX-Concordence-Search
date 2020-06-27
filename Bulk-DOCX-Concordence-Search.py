import nltk
from nltk.tokenize import word_tokenize
from nltk import Text 
from docx import *
nltk.download('gutenberg')
import docx2txt
import zipfile
import os

while True:
    
    filepath = input('\nPaste the location of the project folder you wish to search: ')

    while True:

        try:
            search = input('\nEnter search word or type "cd" to change folder location: ')
            if search == 'cd':
                break
            
            else:
                print('\nSearching docs for: ' + search)

            text = ''

            for folder_path, folder_names, file_names in os.walk(filepath):
                for i in file_names:
                    if not i.endswith(".docx"):
                         continue
                    path = os.path.join(folder_path, i)
                    text += docx2txt.process(path)


            tokens = word_tokenize(text)    
            textList = Text(tokens)
            terms = textList.concordance(search, width=159, lines = 10000)
        
        except zipfile.BadZipFile as err:
            print('\nYou need to close the file you are searching and try again')
