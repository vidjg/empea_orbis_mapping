# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 10:34:35 2018

@author: sqian
"""

import pandas as pd
from fuzzywuzzy import fuzz



raw = pd.read_csv("C:/Users/sqian/PycharmProjects/Bulk_Download/All_columns-2015.txt",delimiter="|")

to_match = pd.read_csv("C:/Users/sqian/OneDrive - WBG/Documents/06-25 - Orbis Investee Match/data_to_map_python.txt",sep='\t')

def match_name(name, list_names, min_score=0):
    # -1 score incase we don't get any matches
    max_score = -1
    # Returning empty name for no match as well
    max_name = ""
    # Iternating over all names in the other
    for name2 in list_names:
        #Finding fuzzy match score
        score = fuzz.ratio(name, name2)
        # Checking if we are above our threshold and have a better score
        if (score > min_score) & (score > max_score):
            max_name = name2
            max_score = score
    return (max_name, max_score)



def main():
    names_to_match = to_match['investee']
    dict_list = []
    for name in names_to_match:
        match = match_name(name, raw['company_name'], 75)    
        # New dict for storing data
        dict_ = {}
        dict_.update({"player_name" : name})
        dict_.update({"match_name" : match[0]})
        dict_.update({"score" : match[1]})
        dict_list.append(dict_)   
    merge_table = pd.DataFrame(dict_list)
    merge_table.to_csv('merged_table.csv',index=False,sep='\t')