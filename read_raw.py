# -*- coding: utf-8 -*-
"""
Created on Mon Jun 25 10:34:35 2018

@author: sqian
"""

import pandas as pd
from fuzzywuzzy



raw = pd.read_csv("C:/Users/sqian/PycharmProjects/Bulk_Download/All_columns-2015.txt",delimiter="|")

to_match = pd.read_csv("C:/Users/sqian/OneDrive - WBG/Documents/06-25 - Orbis Investee Match/data_to_map_python.txt",sep='\t')
