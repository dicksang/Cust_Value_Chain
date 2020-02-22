# -*- coding: utf-8 -*-
"""
Created on Mon Sep 16 21:36:23 2019

@author: Dick Sang
"""

import os.path

dict_path = 'C:\\Users\\Dick Sang\\Desktop\\5. Data Analytics\\3. PolyU RA\\' \
            '1. Projects\\3. Cust Value Chain Analysis\\2. LIWC\\'
working_path = "C:\\Users\\Dick Sang\\Desktop\\5. Data Analytics\\3. PolyU RA\\" \
               "1. Projects\\3. Cust Value Chain Analysis\\" \
               "1. Apple Podcast_speeches\\import files\\"
export_path = "C:\\Users\\Dick Sang\\Desktop\\5. Data Analytics\\3. PolyU RA\\"\
              "1. Projects\\3. Cust Value Chain Analysis\\"

os.chdir(working_path)
import pandas as pd

import liwc
parse, category_names = liwc.load_token_parser(dict_path + 'Cust_val_chain_keywords.dic')

import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
import string
import re
import numpy as np
from dfply import *
from math import *
import glob

from collections import Counter

# locate the files inside the current folder
file_list =  glob.glob("*.DOC")

# Initialize df
df_freq_summary = pd.DataFrame(columns=[
                                        'File', \
                                        'Functional', \
                                        'Experimental', \
                                        'Symbolic', \
                                        'Cost', \
                                        'Optimistic', \
                                        'Affective', \
                                        'Total_WC'\
                                        ])

df_freq_summary = df_freq_summary.astype({'File': 'str'})
df_freq_summary = df_freq_summary.astype({'Functional': 'int'})
df_freq_summary = df_freq_summary.astype({'Experimental': 'int'})
df_freq_summary = df_freq_summary.astype({'Symbolic': 'int'})
df_freq_summary = df_freq_summary.astype({'Cost': 'int'})
df_freq_summary = df_freq_summary.astype({'Optimistic': 'int'})
df_freq_summary = df_freq_summary.astype({'Affective': 'int'})
df_freq_summary = df_freq_summary.astype({'Total_WC': 'int'})

i = 0

for file in file_list:
        if "~" not in file:
            
            file_name = working_path + file
            short_file_name = file[:-4] # remove .doc
                    
            # open word document
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False
            
            word_doc = word.Documents.Open(file_name)
            text = word_doc.Range().Text
            
            # remove time format data eg.(0:12:34, 00:12:34)
            # *-> repeat at least 0 time
            text2 = re.sub("\[\d*:\d*:\d*\]", "", text)
            # remove text, numeric + word format data (each part many times)
            text3 = re.sub("\[\d*\D*\d*\D*\d*\D*\]", "", text2)
            text4 = re.sub("’", "'", text3)
            
            tokenizer = nltk.tokenize.MWETokenizer()
            tokenizer.add_mwe(("he", "'ll"))
            tokenizer.add_mwe(("i", "'ll"))
            tokenizer.add_mwe(("it", "'ll"))
            tokenizer.add_mwe(("she", "'ll"))
            tokenizer.add_mwe(("they", "'ll"))
            tokenizer.add_mwe(("we", "'ll"))
            tokenizer.add_mwe(("you", "'ll"))
            tokenizer.add_mwe(("wo", "n't"))
            
            word_tokens = word_tokenize(text4.lower())
            word_tokens2 = tokenizer.tokenize(word_tokens)
            
                # remove punctuations
            word_tokens3 = [x for x in word_tokens2 if not re.fullmatch('[' + 
                            string.punctuation + ']+', x)]

            word_counts = Counter([category for token in word_tokens3
                                for category in parse(token)])
    
            word_counts['Total_WC'] = len(text4.split())
    
            #count the number of keywords available
            # freq = nltk.FreqDist(word_tokens3)
            
            df_freq_summary = df_freq_summary.append(word_counts, ignore_index=True)
            df_freq_summary.set_value(i, 'File', short_file_name)
            
            i = i + 1
            
df_freq_summary.to_csv(export_path + '1. speech_summary.csv', sep = ',')