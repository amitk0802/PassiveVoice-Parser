# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#import spacy
import time
import pandas as pd
#from re import A
import en_core_web_sm
nlp = en_core_web_sm.load()
import openpyxl

from nltk import tokenize
import nltk
nltk.download('punkt')

import tkinter as tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

import logging

# function to check the type of sentence
def checkForSentType(inputSentence):

    # running the model on sentence
    getDocFile = nlp(inputSentence)
    word_tag_list = []
    count_agent = []
    count_nsubjpass = []
    count_csubjpass = []
    count_auxpass = []

    getAllWordTags = [(token, token.dep_) for token in getDocFile]

    # checking for 'agent' tag
    for (word_, sublist_) in getAllWordTags:
        if sublist_ in ['agent', 'nsubjpass', 'auxpass', 'csubjpass']:
            word_tag_list.append((word_, sublist_))
    
    for (word_, sublist_) in getAllWordTags:
        if sublist_ in ['agent']:
            count_agent.append((word_, sublist_))

    for (word_, sublist_) in getAllWordTags:
        if sublist_ in ['nsubjpass']:
            count_nsubjpass.append((word_, sublist_))

    for (word_, sublist_) in getAllWordTags:
        if sublist_ in ['csubjpass']:
            count_csubjpass.append((word_, sublist_))

    for (word_, sublist_) in getAllWordTags:
        if sublist_ in ['auxpass']:
            count_auxpass.append((word_, sublist_))

    degree_passive = max([len(count_agent), len(count_nsubjpass), len(count_csubjpass), len(count_auxpass)])

    return degree_passive, word_tag_list

if __name__ == '__main__':
    try:
        tk.Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
        input_file = askopenfilename()
        ext = input_file.split(".")[-1]
        print("Extension : {}".format(ext))
        present_date = time.strftime('%Y%m%d') 
        input_file_name = str(input_file.rsplit('/', 1)[1])
        output_file_name = str(input_file_name).replace('.' + str(ext),' - Output from PassiveVoiceParser.' + str(ext))
        output_file = str(input_file.rsplit('/', 1)[0]) + '/' + str(present_date) + '-' + output_file_name
        
        print("Input File : {}".format(input_file))
        time.sleep(2)
        if ext == 'xlsx':
            data_df = pd.read_excel(input_file, sheet_name = 0, engine = 'openpyxl')
            print('File Read Success!')
        
        elif ext == 'xls':
            data_df = pd.read_excel(input_file, sheet_name = 0, engine = 'xlrd')
            print('File Read Success!')
            
        elif ext == 'csv':
            data_df  = pd.read_csv(input_file)
            print('File Read Success!')
            
        else:
            print('FILE FORMAT NOT SUPPORTED!')
            
        col1 = data_df.columns[0]
        col2 = data_df.columns[1]
        data_df.rename(columns={ data_df.columns[0]: "NHTSACampaignNumber",  data_df.columns[1]: "Chronology"}, inplace = True)
        
        data_df['Sentences'] = ''
        data_df['Passivity'] = ''
        data_df['nSent'] = ''
        data_df['WordsAndTags'] = ''
        print('Running...')
        for i in range(len(data_df)):
            new_sentences = tokenize.sent_tokenize(data_df['Chronology'][i])
            sentences = []
            normal_sentence = []
            for new_i in new_sentences:
                split_sent = new_i.split('')
                for new_j in split_sent:
                    split_sent_last = new_j.split('•')
                    for new_k in split_sent_last:
                        split_sent_next = new_k.split('\no ')
                        for new_1 in split_sent_next:
                            split_sent_1 = new_1.split(' (1) ')
                            for new_2 in split_sent_1:
                                split_sent_2 = new_2.split(' (2) ')
                                for new_3 in split_sent_2:
                                    split_sent_3 = new_3.split(' (3) ')
                                    for new_4 in split_sent_3:
                                        split_sent_4 = new_4.split(' (4) ')
                                        for new_5 in split_sent_4:
                                            split_sent_5 = new_5.split(' (5) ')
                                            for new_6 in split_sent_5:
                                                split_sent_6 = new_6.split(' (6) ')
                                                for new_7 in split_sent_6:
                                                    split_sent_7 = new_7.split(' (7) ')
                                                    for new_8 in split_sent_7:
                                                        split_sent_8 = new_8.split(' (8) ')
                                                        for new_9 in split_sent_8:
                                                            split_sent_9 = new_9.split(' (9) ')
                                                            for new_10 in split_sent_9:
                                                                split_sent_10 = new_10.split(' (10) ') 
                                                                for new_l in split_sent_10:
                                                                    normal_sentence.append(new_l)
    
            for crlf in normal_sentence:
                sentences.append(crlf.replace('\n', ' ').replace('\r', ''))
    
            len_list = []
            for cal_len in sentences:
                if len(cal_len) > 15:
                    len_list.append(cal_len)
    
            sent_sum = len(len_list)
            finalResult = []
            # checking each sentence for its type
            for sentence in sentences:
                if len(sentence) > 15:
                    result, word_tag = checkForSentType(str(sentence))
                    data_df = data_df.append({'NHTSACampaignNumber':data_df['NHTSACampaignNumber'][i], 'Chronology': data_df['Chronology'][i], 'Sentences':sentence, 'Passivity': result, 'nSent': sent_sum, 'WordsAndTags':word_tag}, ignore_index=True)
            
            
    
    
        drop_indexes = data_df[((data_df.Sentences==''))].index
        data_df.drop(drop_indexes, inplace=True)
        data_df['ctrSent'] = data_df.groupby(['NHTSACampaignNumber']).cumcount()+1
        data_df = data_df[['NHTSACampaignNumber', 'Chronology', 'Sentences', 'nSent', 'ctrSent', 'Passivity', 'WordsAndTags']]
        data_df.rename(columns={ "NHTSACampaignNumber" : col1, "Chronology" : col2}, inplace = True)
        data_df.reset_index(drop=True, inplace = True)
        
        for k in range(len(data_df)):
            list_val = data_df['WordsAndTags'][k]
            if  len(list_val) == 0:
                data_df['WordsAndTags'][k] = ''
        
        if (ext == 'xlsx') or (ext == 'xls'):
            data_df.to_excel(output_file, sheet_name = 'Sheet1', index = False)
        elif ext == 'csv':
            data_df.to_csv(output_file, index = False)
        
        print('\nCOMPLETED SUCCESSFULLY !!')
        
    except Exception as e:
        logfile = str(input_file.rsplit('/', 1)[0]) + '/' + str(present_date) + ' - PassiveVoiceParser.log' 
        logging.basicConfig(filename=logfile, 
                            level=logging.DEBUG)
        logger=logging.getLogger(__name__)
        logger.error(e)
        logger.info(input_file)
