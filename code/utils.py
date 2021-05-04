#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  4 23:29:11 2021

@author: caiazzo
"""

def get_choices(choices,v_name,lang='en'):
    df = choices.loc[choices['list_name'] == v_name]
    return {
        "name": df['name'].values.astype(str),
        "label": df[get_label(lang)].values.astype(str)
        }

def get_label(lang):
    # english
    if lang=="en": return 'label::English (en)'

    elif lang=="fr": return 'label::Français (fr)'
    
    elif lang=='es': return 'label::Español (es)'
    
    else: 
        print("Warning: unknown language - using english")
        return 'label::English (en)'


def get_hint(lang):
    # english
    if lang == "en": return 'hint::English (en)'

    elif lang=="fr": return 'hint::Français (fr)'
    
    elif lang=='es': return 'hint::Español (es)'
    
    else:
        print("Warning: unknown language - using english")
        return 'hint::English (en)'
    
    
def get_percent(n,l):
    if l==0: return 0
    else:
        return round(n/l*100,1) 
    

def get_question_text(l):
    q = l
    q = q.replace('{', '')
    q = q.replace('\n', '')
    q = q.replace('}', '')
    q = q.replace('_', '\_')
    q = q.replace('$', '')
    return q

def get_color(v):
    if v<=20: return "color0"
    elif v<=40: return "color1"
    elif v<=60: return "color2"
    elif v <= 80: return "color3"
    else: return "color4"


