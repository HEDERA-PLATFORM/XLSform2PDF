#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  4 23:05:17 2021

@author: caiazzo
"""



class InputParameters:
    def __init__(self,input_dict):
        self.verbose = input_dict['verbose'] if 'verbose' in input_dict.keys() else 0
        
        # REQUIRED
        self.filepath = input_dict['filepath'] 
        self.outputTexFile = input_dict['outputTexFile']
        # survey
        self.survey_name = input_dict['survey_name']
        self.section_groups = input_dict['section_groups']
        self.lang = input_dict['lang']
        
        # OPTIONAL
        # layout
        self.logo = input_dict['logo'] if 'logo' in input_dict.keys() else None
        self.fontFamily = input_dict['fontFamily'] if 'fontFamily' in input_dict.keys() else None
        self.fontpath = input_dict['fontpath'] if 'fontpath' in input_dict.keys() else None
        #fontpath = '/Library/Fonts/JosefinSans-Regular.ttf'
        
        
        # optional data file
        self.submissions_name = input_dict['submissions_name'] if 'submissions_name' in input_dict.keys() else None
        self.date_key = input_dict['date_key'] if 'date_key' in input_dict.keys() else None
        self.figdir = input_dict['figdir'] if 'figdir' in input_dict.keys() else None
        
        # do not change
        self.GITHUB_URL = "https://github.com/HEDERA-PLATFORM/XLSform2PDF"
        self.COLOR_SCALE_RED = ["#F5B7B1","#EC7063","#E74C3C","#B03A2E","#943126"] #from https://htmlcolorcodes.com/
        self.COLOR_SCALE_BLUE = ["#D4E6F1","#7FB3D5","#2980B9","#1F618D","#154360"]
        self.COLOR_SCALE_ORANGE = ["#FAE5D3","#EDBB99","#DC7633","#BA4A00","#873600"]
###############################################################################