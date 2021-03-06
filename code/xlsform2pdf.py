#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Apr 23 16:50:58 2021

@author: caiazzo
"""


import pandas as pd
import numpy as np
from datetime import datetime
from collections import Counter
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import json
import sys
import nltk
import re
from nltk.corpus import stopwords
nltk.download('stopwords')
from nltk.tokenize import word_tokenize
import os
nltk.download('punkt')


from input_parameters import InputParameters
from utils import (
    get_choices, get_label, get_hint, get_percent, get_question_text, get_color
    )


# read input file (.json)    
#input_file = sys.argv[1]

input_file = "/Users/caiazzo/HEDERA/CODES/XLSform2PDF/demo/i_fansoto.json"
with open(input_file) as f:
    input_dict = json.load(f)

io = InputParameters(input_dict)

if io.verbose>0:
    print(input_dict)

#[]
# fontpath = io.fontpath
# fm.fontManager.addfont(fontpath)
# prop = fm.FontProperties(fname=fontpath)
# matplotlib.rcParams['font.family'] = prop.get_name()
# plt.rcParams['font.family'] = prop.get_name()
# plt.rcParams['font.size'] = 18       



# ##############################################################################


# read survey
survey = pd.read_excel(io.survey_name, sheet_name = "survey")
choices = pd.read_excel(io.survey_name, sheet_name = "choices")
settings = pd.read_excel(io.survey_name, sheet_name = "settings")
s_version = settings['version'].values[0]
s_name = settings['name'].values[0]
s_title = settings['form_title'].values[0]



if not io.submissions_name==None:
    submissions = pd.read_csv(io.submissions_name)
    last_date_value = submissions[io.date_key][0][:10]
    date_object = datetime.strptime(last_date_value, io.date_format)
    last_date = date_object.strftime("%d %B, %Y")
    
    new_names = []
    for c in submissions.columns:
        ic = c.rfind('-')
        new_names.append(c[ic+1:])
    submissions.columns = new_names

# prepare directory
if not os.path.exists(io.outputDirectory):
    # create directory
    os.system('mkdir ' + io.outputDirectory)
    os.system('cp assets/' + io.logo  + ' ' + io.outputDirectory)
    os.system('cp assets/' + io.fontFile + ' ' + io.outputDirectory)
       
texFile = io.outputDirectory + "/main.tex"
print(" *** writing survey to file " + io.outputDirectory + "/main.tex *** ")
fnew = open(texFile,'w')

## write header
fnew.write("\documentclass[11.5pt, a4paper]{scrartcl}\n")
fnew.write("\\usepackage[english]{babel}\n")

fnew.write("\\usepackage{fontspec}\n")

fnew.write('\\usepackage[top=35pt, left=2.25cm, right=2.2cm]{geometry}\n')
fnew.write('\\usepackage{setspace}\n')
fnew.write('\\usepackage{fancyhdr}\n')
fnew.write('\\pagestyle{fancy}\n')
fnew.write('\\fancyhf{moresize}\n')
fnew.write('\\usepackage{hyperref}\n')
fnew.write('\\usepackage{float}\n')

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
fnew.write('\\fancyhead[R]{\\small ' + settings['form_title'].values[0].replace('_',' ') + '}\n')
fnew.write('\\fancyfoot[R] { {\\textcolor{white}\\thepage}}\n')

fnew.write('\\fancypagestyle{plain}{%\n')
fnew.write('\\fancyhf{}\n')
fnew.write('\\fancyfoot[R] {\Large \\thepage}\n')
fnew.write('}\n')


for c in io.COLOR_SCALE_ORANGE:
    fnew.write('\\definecolor{color'+str(io.COLOR_SCALE_ORANGE.index(c)) + '}{HTML}{' + c[1:] + '}\n')

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
if not io.logo == None:
    fnew.write('\\node (mytext) at (4.2,0.53) { {\\textcolor{white}{\\textsf{POWERED BY}}}};\n');
    fnew.write('\\node (myfirstpic) at (6,0.33) {\\includegraphics[height=.7cm]{' + io.logo + '}};\n');

fnew.write('\\end{tikzpicture}}\n')
fnew.write('}\n')


if not io.fontFile==None:
    fnew.write('\\setmainfont{[' + io.fontFile + ']}\n')

fnew.write('\\setcounter{secnumdepth}{1}\n')
fnew.write('\\parindent 0pt\n')

fnew.write('\\tableofcontents\n')

fnew.write('\n')

# copyright & info minipage
fnew.write('\\vspace*{\\fill}\n')
fnew.write('\\begin{minipage}[t]{0.7\\textwidth}\n')
fnew.write('{\\small\n')
fnew.write('\\begin{flushleft}\n')
fnew.write('HEDERA XLSForm$\\_$Explore v1.0 (May 2021) \\\\[0.2em]\n')
fnew.write('XLSForm$\\_$Explore is an open source project to\n')
fnew.write('automatically create a codebook from XLS Forms (view on \\href{' + io.GITHUB_URL + '}{github})\\\\\n')
fnew.write('Survey: ' + s_title.replace('_','$\\_$') + '\\\\\n')
fnew.write('ID: ' + s_name.replace('_','$\\_$') + ' -- version ' + str(s_version) +'\\\\\n')
if not io.submissions_name == None:
    fnew.write('Data file: ' + io.submissions_name.replace('_','$\\_$') + ' (last update on: ' + last_date + ') \\\\\n')
fnew.write('?? Copyright \\href{https://hedera.online}{HEDERA Sustainable Solutions}\n')
fnew.write('\\end{flushleft}\n')
fnew.write('}')
fnew.write('\\end{minipage}\n')


# fnew.write('\\newpage\n')


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
        if variable_name in io.section_groups:
            first_name = variable_name
            surveyStart = True
            q = row[get_label(io.lang)]
            fnew.write('\\newpage')
            fnew.write('\\section{'+q + '}\n')
            
    
    if surveyStart:
        
        if not io.figdir == None:
            
            # experimental: draw and include graphs
            ###################################################################
            if variable_type == "integer":
            ###################################################################
                if len(submissions) > submissions[row['name']].isna().sum():
                    
                    fig,ax = plt.subplots(figsize=(12,8))
                    plt.hist(submissions[row['name']])
                    plt.xlabel(row['name'])
                    plt.tight_layout()
                    plt.savefig(io.figdir + row['name'] + ".png")
            
            
            ###################################################################
            if variable_type.split()[0] == 'select_one':
            ###################################################################
                fig,ax = plt.subplots(figsize=(12,8))
                counts = submissions[row['name']].value_counts().to_dict()
                plt.pie([float(v) for v in counts.values()], labels=[k for k in counts],autopct=None)
                plt.tight_layout()
                plt.savefig(io.figdir + row['name'] + ".png")
                
            
        if (variable_type.split()[0] == 'select_one' or 
            variable_type.split()[0] == 'select_multiple' or 
            variable_type in relevant_types):
            
            if variable_type.split()[0] == 'select_one':
                vtype = "single choice"
            elif variable_type.split()[0] == 'select_multiple':
                vtype = "multiple choice"
            else:
                vtype = variable_type
            
                
            q_label = get_question_text(row[get_label(io.lang)])
            fnew.write('\\paragraph{'+q_label.replace('*','') +'}\n')
            if type(row[get_hint(io.lang)]) == str:
                h = get_question_text(row[get_hint(io.lang)])
                fnew.write('\\ \\\ {\\small ' + h + '}\n')
    
            fnew.write('\\  \\\\')
            
            fnew.write('Variable name: \\texttt{' + variable_name.replace('_','\_') + '}')
            if required:
                fnew.write('\\hfill\\colorbox{red}{\\small{\\textcolor{white}{required}}}\\\\\n ')
            else:
                fnew.write('\\\\\n')
            
                
        
            fnew.write('Type: '+ vtype + '\\\\\n')
            
            
        
        #######################################################################
        if (variable_type.split()[0] == 'select_one' or 
            variable_type.split()[0] == 'select_multiple'):
        #######################################################################
            
            second_name = row['name']
            choices_list = get_choices(choices, variable_type.split()[1],lang=io.lang)

            
            ## About the results in the submissions table:
            if not io.submissions_name == None:
                
                col = row['name'].rstrip()
                total = len(submissions) - submissions[col].isna().sum()
                if 'nan' in submissions[col].value_counts():
                    total = total - submissions[col].value_counts()['nan']
                
                if total>0:
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
                        l = l.replace('%','\\%')
    
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

        #######################################################################
        if variable_type == "integer" or variable_type == "decimal" or (variable_type == "text" and row['appearance']=="numbers") :
        #######################################################################
            col = row['name'].rstrip()
            if variable_type == "text" and row['appearance']=="numbers":
                submissions[col] = pd.to_numeric(submissions[col],errors='coerce')


            data = submissions[[col]].astype(float).dropna()
            data = data[data[col] != 888]
            data = data[data[col] != 8888]
            
            min_value = data[col].min()
            max_value = data[col].max()
            mean_value = np.mean(data[col])
            variance = np.var(data[col])
            
            n_steps = 7
            ns = []
            r0 = min_value
            for k in range(0,n_steps-1):
                r0 += (max_value - min_value)/n_steps
                nsum = sum(ns)
                ns.append(len(data.loc[data[col]<=r0]) - nsum)
            ns.append(len(data)-sum(ns))
            
            
            
            nb_answers = len(data)
            
            fnew.write('\\\\Total number of answers: ' + str(nb_answers) + ' (' +
                       str(get_percent(nb_answers, len(submissions))) + '\\%)\n')
            fnew.write('\\\\[0.2em]\n')

            if nb_answers > 0:

                fnew.write(' \\begin{tabular}{p{4cm}|p{8cm}}\n')
                fnew.write('Minumum &')
                fnew.write('{}'.format(round(min_value,2)))
                fnew.write(' \\\\\n')
                fnew.write('\\hline\n')

                fnew.write('\\cellcolor{mygray} Maximum value & \\cellcolor{mygray}')
                fnew.write('{}'.format(round(max_value,2)))
                fnew.write(' \\\\\n')
                fnew.write('\\hline\n')

                fnew.write('Mean &')
                fnew.write('{}'.format(round(mean_value,2)))
                fnew.write(' \\\\\n')
                fnew.write('\\hline\n')

                fnew.write('\\cellcolor{mygray} Variance & \\cellcolor{mygray}')
                fnew.write('{}'.format(round(variance,2)))
                fnew.write(' \\\\\n')
                fnew.write('\\hline\n')

                fnew.write('\\end{tabular}\n')

                fnew.write('\\\\[0.5em]\n')
                fnew.write('Distribution\\\\\n')
                
                
                fnew.write(' \\begin{tabular}{')
                for k in range(0,n_steps):
                    fnew.write('p{1.9cm}')
                    if k<n_steps-1:
                        fnew.write('|')
                fnew.write('}\n')
                
                for k in range(0,n_steps):
                    fnew.write('{}'.format( round(100/n_steps*(k+1))  ))
                    fnew.write('\\%')
                    if k<n_steps-1:
                        fnew.write('&')
                fnew.write('\\\\\n')
                #fnew.write('25\\% & 50\\% & 75\\% & 100\\% \\\\\n')
                
                for k in range(0,n_steps):
                    value = round(ns[k]/len(data)*100,1)
                    fnew.write('\\cellcolor{' + get_color(value) +'}')
                    fnew.write('{}'.format(ns[k]))
                    fnew.write('(' + str(value) + '\%)')
                    if k<n_steps-1:
                        fnew.write('&')
                
                fnew.write('\\\\\n')
                fnew.write('\\end{tabular}\n')
                
                
                if not io.OutputPlotsFolder == None:
                    fig = plt.figure()
                    title = 'Boxplot of {}'.format(col)
                    ax = submissions[col].plot.box(title = title)
                    ax.plot()
    
                    file_path = io.OutputPlotsFolder + '/graph_' + '{}'.format(col) + '.png'
    
                    fig.savefig(file_path)
                    fig = plt.close()
    
                    fnew.write('\\begin{figure}[H]\n')
                    fnew.write('\\centering\n')
                    fnew.write('\\includegraphics[scale=0.5]{')
                    fnew.write('{}'.format(file_path))
                    fnew.write('}\n')
                    fnew.write('\\end{figure}\n')

        #######################################################################
        if io.textAnalysis:
            if variable_type == "text" and not row['appearance']=="numbers":
                col = row['name'].rstrip()
                nb_answers = len(submissions[col].dropna())
                s = submissions[col].dropna()
    
                text = []
                if len(s)>0:
                    for index, value in s.items():
                        #for word in value:
                         text.append(value.lower())
                    strings = ' '.join(text)
                    strings = re.sub(r'[^\w\s]', '', strings)
        
                    text_tokens = word_tokenize(strings)
                    tokens_without_sw = [word for word in text_tokens if not word in stopwords.words()]
                    counter = Counter(tokens_without_sw)
                    most_occur = counter.most_common(8)
                    #print('most_occur', most_occur)
        
                    fnew.write('\\\\Total number of answers: ' + str(total) + ' (' +
                               str(get_percent(total, len(submissions))) + '\\%)\n')
                    fnew.write('\\\\[0.2em]')
        
                    if not most_occur == []:
                        fnew.write('\\begin{table}[H]\n')
                        fnew.write(' \\begin{tabular}{p{4cm}|p{8cm}}\n')
                        fnew.write('Word & Number of occurrences ')
                        fnew.write(' \\\\\n')
                        fnew.write('\\hline\n')
                        for i in range(len(most_occur)):
                            if i%2 == 1:
                                fnew.write('{}'.format(most_occur[i][0]))
                                fnew.write('&')
                                fnew.write('{}'.format(most_occur[i][1]))
                                fnew.write('\\\\\n')
                                fnew.write('\\hline\n')
                            else:
                                fnew.write('\\cellcolor{mygray}')
                                fnew.write('{}'.format(most_occur[i][0]))
                                fnew.write('&')
                                fnew.write('\\cellcolor{mygray}')
                                fnew.write('{}'.format(most_occur[i][1]))
                                fnew.write('\\\\\n')
                                fnew.write('\\hline\n')
        
        
                        fnew.write('\\end{tabular}\n')
                        fnew.write('\\caption{\\label{tab:table-name} Most used words for this answer}\n')
                        fnew.write('\\end{table}\n')






        #######################################################################
        if (variable_type == 'range'):
        #######################################################################
            if not io.submissions_name == None:
                # write the range results
                col = row['name'].rstrip()
                total = len(submissions) - submissions[col].isna().sum()
                values = np.unique(submissions[col].values)
                # remove nan
                values = values[~np.isnan(values)]
                
                if total>0:
                    fnew.write('\\\\Total number of answers: ' + str(total) + ' (' + 
                               str(get_percent(total,len(submissions))) + '\\%)\n')
                    fnew.write('\\\\[0.2em]')
                    
    
                    fnew.write(' \\begin{tabular}{p{4cm}|p{4cm}}\n')
                    fnew.write('Value & Answers \\\\\n')
                    fnew.write('\\hline\n')
                    
                    for v in values:
                        c_name = v.astype(str)
                        f = sum(submissions[col]==v)
                        value = get_percent(f, total)
                        cell_color = get_color(value)
                        if k%2==1:
                            #fnew.write('\\cellcolor{mygray} '+ n+ ' & \\cellcolor{mygray}' + l + ' & \\cellcolor{mygray}' + str(f) + ' (' + get_percent(f,total) + '\%)\\\\\n')
                            fnew.write('\\cellcolor{mygray} '+ c_name + ' & \\cellcolor{' + cell_color + '}' + str(f) + ' (' + str(value) + '\%)\\\\\n')
                        else:
                            fnew.write( c_name + ' &  \\cellcolor{' + cell_color + '}' + str(f) + ' (' + str(value) + '\%)\\\\\n')
    
                    fnew.write('\\end{tabular}\n')




    



fnew.write('\\end{document}')
fnew.close()
