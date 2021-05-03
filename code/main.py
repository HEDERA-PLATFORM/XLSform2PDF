#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr 23 16:50:58 2021

@author: caiazzo
"""


import pandas as pd
import numpy as np
import math
from datetime import datetime
import matplotlib
import matplotlib.pyplot as plt

# color scales from https://htmlcolorcodes.com/
COLOR_SCALE_RED = ["#F5B7B1","#EC7063","#E74C3C","#B03A2E","#943126"]
COLOR_SCALE_BLUE = ["#D4E6F1","#7FB3D5","#2980B9","#1F618D","#154360"]
COLOR_SCALE_ORANGE = ["#FAE5D3","#EDBB99","#DC7633","#BA4A00","#873600"]

def get_choices(choices,v_name,lang='en'):
    df = choices.loc[choices['list_name'] == v_name]
    return {
        "name": df['name'].values.astype(str),
        "label": df[get_label(lang)].values.astype(str)
        }

def get_label(lang):
    # english
    if lang=="en":
        return 'label::English (en)'

    elif lang=="fr":
        return 'label::Français (fr)'
    
    elif lang=='es':
        return 'label::Español (es)'
    
    else:
        print("Warning: unknown language - using english")
        return 'label::English (en)'


def get_hint(lang):
    # english
    if lang == "en":
        return 'hint::English (en)'

    elif lang=="fr":
        return 'hint::Français (fr)'
    
    elif lang=='es':
        return 'hint::Español (es)'
    
    else:
        print("Warning: unknown language - using english")
        return 'hint::English (en)'
    
    
def get_percent(n,l):
    if l == 0:
        return 0
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
    if v<=20:
        return "color0"
    elif v<=40:
        return "color1"
    elif v<=60:
        return "color2"
    elif v <= 80:
        return "color3"
    else:
        return "color4"
        
    
###############################################################################
#filepath = "./input/"
filepath = "/Users/caiazzo/HEDERA/CODES/XLSform2PDF/"        
    
#outputTexFile = "./output/survey.tex"
outputTexFile = "/Users/caiazzo/HEDERA/CODES/XLSform2PDF/output/zni/zni.tex" 
survey_name = "/Users/caiazzo/HEDERA/NextCloud/IMPACT-R_Project/Activities/Data_Collection/Parallel Projects/ZNI/Surveys/zni_form.xlsx"
figdir = None 
# if True: use special fonts (! must be installed locally !)
#fontFamily = "Josefin Sans"
fontFamily = "Josefin Sans"
# submissions_name = None => no data file attached
submissions_name = None #filepath + "XLSFormExplore/svf.xlsx"
logo = "HEDERA.png"
lang= "es"
# specify the groups that will be separated as sections
#section_groups = ['identification','household','electricity','cooking','jmp_wash','fies','dietary_score']
section_groups = ['datos_generales','beneficiario','mtf','equipos','instalacion']
GITHUB_URL = "https://github.com/HEDERA-PLATFORM/XLSform2PDF"
###############################################################################








### TODO: can we also add - analogously - subsection_groups? 
# attention: it might get complicated with names of groups


#these lines are needed to add the Josefin Sans font
#@attention: the font path might vary depending on the machine
import matplotlib.font_manager as fm
fontpath = '/Library/Fonts/JosefinSans-Regular.ttf'
fm.fontManager.addfont(fontpath)
prop = fm.FontProperties(fname=fontpath)
matplotlib.rcParams['font.family'] = prop.get_name()
plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['font.size'] = 18

###############################################################################

if not fontFamily == None:
    print(" -- creating latex with font family:", fontFamily, " : it should be already installed in order to compile")

# read survey
survey = pd.read_excel(survey_name, sheet_name = "survey")
choices = pd.read_excel(survey_name, sheet_name = "choices")
settings = pd.read_excel(survey_name, sheet_name = "settings")
s_version = settings['version'].values[0]
s_name = settings['name'].values[0]
s_title = settings['form_title'].values[0]

if not submissions_name==None:
    submissions = pd.read_excel(submissions_name, sheet_name = 'SheetJS')
    last_date_value = submissions['Submitted on'][0][:10]
    date_object = datetime.strptime(last_date_value, "%Y-%m-%d")
    last_date = date_object.strftime("%d %B, %Y")
    
    new_names = []
    for c in submissions.columns:
        ic = c.rfind('-')
        new_names.append(c[ic+1:])
    submissions.columns = new_names



print(" *** writing survey to file " + outputTexFile + " *** ")
fnew = open(outputTexFile,'w')

### TODO: check if we can remove some lines
## write header
fnew.write("\documentclass[11.5pt, a4paper]{scrartcl}\n")
fnew.write("\\usepackage[english]{babel}\n")

if not fontFamily == None:
    #fnew.write("\\usepackage{josefin}\n")
    #fnew.write("\\usepackage[sfdefault]{josefin}\n")
    fnew.write("\\usepackage{fontspec}\n")
fnew.write('\\usepackage[top=35pt, left=2.25cm, right=2.2cm]{geometry}\n')
fnew.write('\\usepackage{setspace}\n')
fnew.write('\\usepackage{fancyhdr}\n')
fnew.write('\\pagestyle{fancy}\n')
fnew.write('\\fancyhf{moresize}\n')
fnew.write('\\usepackage{hyperref}\n')

fnew.write('\\usepackage{multicol}\n')
fnew.write('\\usepackage[dvipsnames,table]{xcolor}\n')
fnew.write('\\usepackage{soul}\n')
fnew.write('\\usepackage{tikz}\n')
fnew.write('\\usetikzlibrary{calc}\n')
fnew.write('\\usepackage[placement=bottom,scale=1,opacity=1]{background}\n')


fnew.write('\\setlength{\\headheight}{40pt}\n')
fnew.write('\\setlength{\\footskip}{60pt}\n')

fnew.write('\\usepackage{eurosym}\n')
fnew.write('\\usepackage{booktabs}\n')
fnew.write('\\setlength{\\heavyrulewidth}{1.5pt}\n')
fnew.write('\\setlength{\\abovetopsep}{4pt}\n')


fnew.write('\\fancyhf{}\n')
fnew.write('\\fancyhead[R]{\\small ' + settings['form_title'].values[0] + '}\n')
fnew.write('\\fancyfoot[R] { {\\textcolor{white}\\thepage}}\n')

fnew.write('\\fancypagestyle{plain}{%\n')
fnew.write('\\fancyhf{}\n')
fnew.write('\\fancyfoot[R] {\Large \\thepage}\n')
fnew.write('}\n')


for c in COLOR_SCALE_ORANGE:
    fnew.write('\\definecolor{color'+str(COLOR_SCALE_ORANGE.index(c)) + '}{HTML}{' + c[1:] + '}\n')

fnew.write('\\definecolor{mygreen1}{HTML}{6CAE33}\n')
fnew.write('\\definecolor{mygray}{HTML}{EAECEE}\n')
fnew.write('\\hypersetup{\n')
fnew.write('    colorlinks,\n')
fnew.write('    linkcolor={hederablue},\n')
fnew.write('    citecolor={mygreen1!50!black},\n')
fnew.write('    urlcolor={mygreen1!80!black}\n')
fnew.write('}\n')


# BEGIN DOCUMENT
fnew.write('\\begin{document}\n')
fnew.write('\\definecolor{hederablue}{HTML}{283747}\n')
fnew.write('\\backgroundsetup{contents={%\n')
fnew.write('\\begin{tikzpicture}\n')
fnew.write('\\shade[left color=hederablue, middle color = hederablue, right color = hederablue] (0,1.3) rectangle (23,0);\n')
if not logo == None:
    fnew.write('\\node (mytext) at (4.2,0.53) { {\\textcolor{white}{\\textsf{POWERED BY}}}};\n');
    fnew.write('\\node (myfirstpic) at (6,0.33) {\\includegraphics[height=.7cm]{' + logo + '}};\n');

fnew.write('\\end{tikzpicture}}\n')
fnew.write('}\n')


if not fontFamily==None:
    fnew.write('\\setmainfont{' + fontFamily + '}\n')

fnew.write('\\setcounter{secnumdepth}{1}\n')
fnew.write('\\parindent 0pt\n')

# TABLE OF CONTENTS (use "2" for printing in two columns, for long surveys)
#fnew.write('\\begin{multicols}{1}\n')
#fnew.write('{\\tiny\n')
fnew.write('\\tableofcontents\n')

fnew.write('\n')

# copyright & info minipage
fnew.write('\\vspace*{\\fill}\n')
fnew.write('\\begin{minipage}[t]{0.7\\textwidth}\n')
fnew.write('{\\small\n')
fnew.write('\\begin{flushleft}\n')
fnew.write('HEDERA XLSForm$\\_$Explore v1.0 (May 2021) \\\\[0.2em]\n')
fnew.write('XLSForm$\\_$Explore is an open source project to\n')
fnew.write('automatically create a codebook from XLS Forms (view on \\href{' + GITHUB_URL + '}{github})\\\\\n')
fnew.write('Survey: ' + s_title.replace('_','$\\_$') + '--')
fnew.write('ID: ' + s_name.replace('_','$\\_$') + ' -- version ' + str(s_version) +'\\\\\n')
if not submissions_name == None:
    fnew.write('Data file: ' + submissions_name.replace('_','$\\_$') + ' (last update on: ' + last_date + ') \\\\\n')
fnew.write('© Copyright \\href{https://hedera.online}{HEDERA Sustainable Solutions}\n')
fnew.write('\\end{flushleft}\n')
fnew.write('}')
fnew.write('\\end{minipage}\n')

#if not submissions_name==None:
#    fnew.write('\n')
#    fnew.write('\\ For each question, this file includes some information about the result of this survey which had its last submissions on the ' + date + '. \n')


#fnew.write('}\n')
#fnew.write('\\end{multicols}\n')
fnew.write('\\newpage\n')


# select which variables / questions are we going to write
relevant_types = ['integer', 'text', 'decimal', 'range', 'image','audio']
surveyStart = False


for index, row in survey.iterrows():
    
    variable_type = row['type']
    variable_name = row['name'].rstrip() if not(variable_type)=="end group" else None
    
    
    
    if row['required'] == 'yes':
        required = True
    else:
        required = False
    
    if variable_type.split()[0] == 'begin':
        if variable_name in section_groups:
            first_name = variable_name
            surveyStart = True
            q = row[get_label(lang)]
            fnew.write('\\newpage')
            fnew.write('\\section{'+q + '}\n')
            
    plt.rcParams["font.family"] = fontFamily
    if surveyStart:
        
        if not figdir == None:
            # experimental: draw and include graphs
            if variable_type=="integer":
                if len(submissions) > submissions[row['name']].isna().sum():
                    
                    fig,ax = plt.subplots(figsize=(12,8))
                    plt.hist(submissions[row['name']])
                    plt.xlabel(row['name'])
                    plt.tight_layout()
                    plt.savefig(figdir + row['name'] + ".png")
            
            
            if variable_type.split()[0] == 'select_one':
                fig,ax = plt.subplots(figsize=(12,8))
                counts = submissions[row['name']].value_counts().to_dict()
                plt.pie([float(v) for v in counts.values()], labels=[k for k in counts],autopct=None)
                plt.tight_layout()
                plt.savefig(figdir + row['name'] + ".png")
                
            
        if variable_type.split()[0] == 'select_one' or variable_type.split()[0] == 'select_multiple' or variable_type in relevant_types:
            
            if variable_type.split()[0] == 'select_one':
                vtype = "single choice"
            elif variable_type.split()[0] == 'select_multiple':
                vtype = "multiple choice"
            else:
                vtype = variable_type
            
            q_label = get_question_text(row[get_label(lang)])
            fnew.write('\\paragraph{'+q_label.replace('*','') +'}\n')
            if type(row[get_hint(lang)]) == str:
                h = get_question_text(row[get_hint(lang)])
                fnew.write('\\ \\\ {\\small ' + h + '}\n')
    
            fnew.write('\\  \\\\')
            
            fnew.write('Variable name: \\texttt{' + variable_name.replace('_','\_') + '}')
            if required:
                fnew.write('\\hfill\\colorbox{red}{\\small{\\textcolor{white}{required}}}\\\\\n ')
            else:
                fnew.write('\\\\\n')
            
                
        
            fnew.write('Type: '+ vtype + '\\\\\n')
            
            
        
        
        if variable_type.split()[0] == 'select_one' or variable_type.split()[0] == 'select_multiple':
            
            second_name = row['name']
            choices_list = get_choices(choices, variable_type.split()[1],lang=lang)

            
            ## About the results in the submissions table:
            if not submissions_name == None:
                
                col = row['name'].rstrip()
                
                
                
                total = len(submissions) - submissions[col].isna().sum()
                if 'nan' in submissions[col].value_counts():
                    total = total - submissions[col].value_counts()['nan']
                
                fnew.write('\\\\Total number of answers: ' + str(total) + ' (' + str(get_percent(total,len(submissions))) + '\\%)\n')
                fnew.write('\\\\[0.2em]')
                

                fnew.write(' \\begin{tabular}{p{4cm}|p{8cm}|p{3cm}}\n')
                fnew.write('Choice code & Label & Answers \\\\\n')
                fnew.write('\\hline\n')
                
                for k in range(0, len(choices_list['name'])):
                    n = str(choices_list['name'][k])
                    n = n.replace('_', '\_')
                    c_name = choices_list['name'][k]

                    l = str(choices_list['label'][k])
                    l = l.replace('&','\\&')

                    if variable_type.split()[0] == 'select_one':
                        dict_select_one = submissions[col].astype(str).value_counts().to_dict()
                        f = dict_select_one[c_name] if c_name in dict_select_one else 0
                        
                    
                    else:
                        
                        submissions[col] = submissions[col].astype(str)
                        f = 0
                        for val in submissions[col]:
                            if c_name in val.split(" "):
                                f += 1
                        

                    value = get_percent(f, total)
                    cell_color = get_color(value)
                    if k%2==1:
                        #fnew.write('\\cellcolor{mygray} '+ n+ ' & \\cellcolor{mygray}' + l + ' & \\cellcolor{mygray}' + str(f) + ' (' + get_percent(f,total) + '\%)\\\\\n')
                        fnew.write('\\cellcolor{mygray} '+ n+ ' & \\cellcolor{mygray}' + l + ' & \\cellcolor{' + cell_color + '}' + str(f) + ' (' + str(value) + '\%)\\\\\n')
                    else:
                        fnew.write( n+ ' & ' + l +  '& \\cellcolor{' + cell_color + '}' + str(f) + ' (' + str(value) + '\%)\\\\\n')

                fnew.write('\\end{tabular}\n')

            else:

                fnew.write(' \\begin{tabular}{p{4cm}|p{11cm}}\n')
                fnew.write('Choice code & Label \\\\\n')
                fnew.write('\\hline\n')
                for k in range(0, len(choices_list['name'])):
                    n = str(choices_list['name'][k])
                    n = n.replace('_', '\_')

                    l = str(choices_list['label'][k])
                    l = l.replace('&', '\\&')

                    if k%2==1:
                        fnew.write('\\cellcolor{mygray} '+ n + ' & \\cellcolor{mygray}' + l + '\\\\\n')
                    else:
                        fnew.write( n+ ' & ' + l + '\\\\\n')

                fnew.write('\\end{tabular}\n')


        
            
            

    
    
    


fnew.write('\\end{document}')
fnew.close()
