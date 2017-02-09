__author__ = 'heath'
import sys, os, csv, jinja2
import pandas as pd
import numpy as np
import sqlite3

# Location of file
Location = 'data/results_grid.xlsx'

# Parse a specific sheet
df = pd.read_excel(Location,0,converters={'ID Number':str})


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

for i, row in df.iterrows():
    template = report_renderer.get_template('templates/report_template.tex')
    name = row['Student Name']
    cwid = str(row['ID Number'])
    exam = 'Exam I'
    keyname = row['Key Name']
    score = row['Score']
    correct = row['# Correct']
    question = str(i)
    blank = row['Blank Count']
    uresp = pd.DataFrame(row['Q 1':]).to_latex(header=False)
    filename = str(name).replace(" ", "_") + '_results.tex'
    folder = 'output'
    outpath = os.path.join(folder, filename)
    outfile = open(outpath, 'w')
    outfile.write(template.render(name=name, cwid=cwid, exam=exam, keyname=keyname, score=score, correct=correct, question=question, uresp=uresp))
    outfile.close()
    os.system("pdflatex -output-directory=" + folder + " " + outpath + '>/dev/null')
os.system("rm "+ folder +"/*.tex")
os.system("rm "+ folder +"/*.aux")
os.system("rm "+ folder +"/*.log")
print('Grader is finished.')
