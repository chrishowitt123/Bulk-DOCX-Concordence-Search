import nltk
from nltk.tokenize import word_tokenize
from nltk import Text
import glob 
from docx import *
nltk.download('gutenberg')
import docx2txt
import pandas as pd
import glob
import zipfile

while True:
    
    PATH = input('\nPaste the location of the files you wish to search: ')
    PATH = PATH + '\*.docx'

    while True:

        try:
            search = input('\nEnter search word or type "cd" to change folder location: ')
            if search == 'cd':
                break
            
            else:
                print('\nSearching docs for: ' + search)

            text = ''
            for file in glob.glob(PATH):
                text += docx2txt.process(file)


            tokens = word_tokenize(text)    
            textList = Text(tokens)
            terms = textList.concordance(search, width=159, lines = 10000)
        
        except zipfile.BadZipFile as err:
            print('\nYou need to close the file you are searching and try again')


            

