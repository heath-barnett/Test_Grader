__author__ = 'heath'
import sys, os, csv, jinja2, glob
import pandas as pd
import numpy as np
import sqlite3

# Location of file
Location = 'data/Exam_I_Results_Grid.xlsx'

# Parse a specific sheet
df = pd.read_excel(Location,0,converters={'ID Number':str})
keys = df.loc[:1,'Q 1':]
results = df.loc[2:]
results = results.assign(Total=round(results['Short Answer']+results['# Correct']/30*40,2).values)
tavg = round(results['Total'].mean(),2)
tmin = results['Total'].min()
tmax = results['Total'].max()

sam = round(results['Short Answer'].mean(),2)
samin = results['Short Answer'].min()
samax = results['Short Answer'].max()

sam = round(results['Short Answer'].mean(),2)
samin = results['Short Answer'].min()
samax = results['Short Answer'].max()

mcm = round(results['# Correct'].mean(),2)
mcmin = results['# Correct'].min()
mcmax = results['# Correct'].max()
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

for i, row in df.loc[2:].iterrows():
    template = report_renderer.get_template('templates/report_template.tex')
    name = row['Student Name']
    cwid = str(row['ID Number'])
    exam = 'Exam I'
    keyname = row['Key Name']
    sans = row['Short Answer']
    correct = row['# Correct']
    total = format(sans + correct/30*40,'.2f')
    blank = str(row['Blank Count'])
    cid = 'Chemistry 1007'
    df1 = row['Q 1':]
    max = len(df1.index)
    key = row['Key Name']
    if key == 'A':
        keyid = 0
    else:
        keyid = 1
    ans = keys.ix[keyid]
    gradeTable = pd.concat([df1, ans], axis=1)
    gradeTable.columns = ['Response', 'Correct']
    gradeTable = pd.DataFrame(gradeTable).to_latex(bold_rows=True, column_format='rcc')

    filename = str(name).replace(" ", "_") + '_results.tex'
    folder = 'output'
    outpath = os.path.join(folder, filename)
    outfile = open(outpath, 'w')
    outfile.write(template.render(name=name, cid=cid, tavg=tavg, tmin=tmin, tmax=tmax, sam=sam, samin=samin, samax=samax, mcm=mcm, mcmin=mcmin,mcmax=mcmax, max=max,cwid=cwid, total=total, exam=exam, keyname=keyname, sans=sans, blank=blank, correct=correct, gradeTable=gradeTable))
    outfile.close()
    os.system('pdflatex -quiet -output-directory=' + folder + " " + outpath)
    #for OS X / Linux
    #os.system("pdflatex -output-directory=" + folder + " " + outpath)
    if i ==3:
        break
os.chdir(folder)
for filename in glob.glob('*.tex'):
    os.remove( filename )
for filename in glob.glob('*.aux'):
    os.remove( filename )
for filename in glob.glob('*.log'):
    os.remove( filename )
os.chdir('..')


print('Grader is finished.')
