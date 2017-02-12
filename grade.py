__author__ = 'heath'
import sys, os, csv, jinja2, glob, smtplib
import pandas as pd
import numpy as np
import sqlite3
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

def send_mail( send_from, send_to, subject, text, files=[], server="outlook.office365.com", port=587, username='hbarnett@ulm.edu', password='Tenrab77', isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject

    msg.attach( MIMEText(text) )

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(os.path.basename(f)))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if isTls: smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

# Location of file
Location = 'data/Exam_I_Results_Grid.xlsx'
Roster = 'data/roster_chem1007_spring_2017.xlsx'

# Parse a specific sheet
df = pd.read_excel(Location,0,converters={'ID Number':str})
dfroster = pd.read_excel(Roster,0,converters={'ID Number':str})
keys = df.loc[:1,'Q 1':]
df = df.merge(dfroster[['Email address', 'ID Number']], on='ID Number')

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
    df1 = row['Q 1':'Q 30']
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

    filename = str(name).replace(" ", "_") + '_Exam_I_Results.tex'
    email = row['Email address']

    folder = 'output'
    outpath = os.path.join(folder, filename)
    outfile = open(outpath, 'w')
    outfile.write(template.render(name=name, cid=cid, tavg=tavg, tmin=tmin, tmax=tmax, sam=sam, samin=samin, samax=samax, mcm=mcm, mcmin=mcmin,mcmax=mcmax, max=max,cwid=cwid, total=total, exam=exam, keyname=keyname, sans=sans, blank=blank, correct=correct, gradeTable=gradeTable))
    outfile.close()
    os.system('pdflatex -quiet -output-directory=' + folder + " " + outpath)
    #for OS X / Linux
    #os.system("pdflatex -output-directory=" + folder + " " + outpath)
    rpt = filename.replace('.tex', '.pdf')
    rpt_path = os.path.join(folder,rpt)
    #send_mail(send_from='hbarnett@ulm.edu', send_to=email, files=[rpt_path], subject='Exam I Results', text='Please see the attached pdf for a breakdown of Exam I. Inside the document you will find your scores, correct answers and some basic statistics. I will hand back your exam on Monday and we can discuss it in more detail at that time.\n\n Heath Barnett')

os.chdir(folder)
for filename in glob.glob('*.tex'):
    os.remove( filename )
for filename in glob.glob('*.aux'):
    os.remove( filename )
for filename in glob.glob('*.log'):
    os.remove( filename )
os.chdir('..')


print('Grader is finished.')
