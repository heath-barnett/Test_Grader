__author__ = 'heath'
import sys, os, csv, jinja2
import pandas as pd
import numpy as np
import sqlite3

# Location of file
Location = 'data/results_grid.xlsx'

# Parse a specific sheet
df = pd.read_excel(Location,0,converters={'ID Number':str})
keys = df.loc[:1,'Q 1':]
results = df.loc[2:]

# Change the default delimiters used by Jinja such that it won't pick up brackets attached to LaTeX macros.
report_renderer = jinja2.Environment(
    block_start_string='\BLOCK{',
    block_end_string='}',
    variable_start_string='\VAR{',
    variable_end_string='}',
    comment_start_string='\#{',
    comment_end_string='}',
    line_statement_prefix='%{',
    line_comment_prefix='%#',
    trim_blocks=True,
    autoescape=False,
    loader=jinja2.FileSystemLoader(os.path.abspath('.'))
)

for i, row in results.iterrows():
    template = report_renderer.get_template('templates/report_template.tex')
    name = row['Student Name']
    cwid = str(row['ID Number'])
    exam = 'Exam I'
    keyname = row['Key Name']
    score = row['Score']
    correct = row['# Correct']
    blank = row['Blank Count']

    df1 = row['Q 1':]
    key = row['Key Name']
    if key == 'A':
        keyid = 0
    else:
        keyid = 1
    ans = keys.ix[keyid]
    gradeTable = pd.concat([df1, ans], axis=1)
    gradeTable.columns = ['User Response', 'Correct Response']
    gradeTable = pd.DataFrame(gradeTable).to_latex()

    filename = str(name).replace(" ", "_") + '_results.tex'
    folder = 'output'
    outpath = os.path.join(folder, filename)
    outfile = open(outpath, 'w')
    outfile.write(template.render(name=name, cwid=cwid, exam=exam, keyname=keyname, score=score, correct=correct, gradeTable=gradeTable))
    outfile.close()
    os.system('pdflatex -output-directory=' + folder + " " + outpath)
    #os.system("pdflatex -outpu\\t-directory=" + folder + " " + outpath)
    if i ==3:
        break
#os.system("rm "+ folder +"/*.tex")
#os.system("rm "+ folder +"/*.aux")
#os.system("rm "+ folder +"/*.log")
print('Grader is finished.')
