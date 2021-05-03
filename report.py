# -*- coding: utf-8 -*-
## Creating Work Sheet PDF

##pip install reportlab

def makeReport(weekof):
  # Import libraries 
  import pandas as pd
  import numpy as np
  from reportlab.pdfgen.canvas import Canvas
  from reportlab.lib import colors
  from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
  from reportlab.lib.units import inch
  from reportlab.platypus import Paragraph, Frame, Table, Spacer, TableStyle, PageBreak
  from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
  from reportlab.lib.pagesizes import letter, landscape

  excel = pd.ExcelFile(f'/content/drive/My Drive/Work_Sheet_{weekof}.xlsx')
  sheet_names = excel.sheet_names
  sheet_names.remove('All_Employees')
  i = 1
  for string in sheet_names:
    new_string = string.replace("_", " ")
    new_string = new_string.replace("Sheet", " ")
    print(i, new_string)
    i += 1
  numbers = []
  for i in range(1, len(sheet_names)+1):
    numbers.append(i)

  name_dict = {numbers[i]: sheet_names[i] for i in range(len(numbers))}
  
  boolean = True
  while boolean:
    print('Enter number of employee you wish to run a report:')
    number = input()
    number = int(number)
    if number >= (len(sheet_names)+1):
      print('Number entered is out of range. Please re-enter number of employee you wish to run a report:')
      number = input()
    number = int(number)
    name = name_dict[number]
    print('The employee sheet you picked is:', name, '. \n Is this correct? (y/n)')
    boolean = input()
    if boolean == 'y':
      boolean = False
    else:
      boolean = True 

  weekof_str = str(weekof)
  #yyyy-mm-dd
  year = weekof_str[:4]
  month = weekof_str[4:6]
  months = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', 
            '07':'July', '08':'August', '09':'September', '10':'October', '11':'November', '12':'December'}
  month = months.get(f'{month}')
  day = weekof_str[6:]
  date = f'{month} {day}, {year}'

  try:
    df = excel.parse(f'{name}')
  except:
    print('This the list of employees that have a sheet:')
    print(sheet_names)
    return 'Incorrect employee name spelling or employee does not have a sheet this week.'
  
  df = df.fillna('-')
  df = df[['LT#', 'Project', 'Board', 'Tape', 'Twelve', 'Ten', 'Nine', 'Eight', 
           'Prepwork','SQF', 'Rate', 'Total','Notes']]	
  df = df.rename(columns={'Twelve':'12\'', 'Ten':'10\'', 'Nine':'9\'', 'Eight':'8\''})

  # Total prior to extras and deductions
  total_ = df.iloc[-12]['Total']

  extras = df.loc[(df['LT#'] == 'Extras:') | (df['LT#'] == 'Total Extras:')]
  extras = extras[['Project', 'Total']]
  extras = extras.loc[extras.Project != 'Enter Text Here']
  extras.Total = extras.Total.astype(float)
  append_Dict = {'Project':'Total Extras:'}
  extras = extras.append(append_Dict, ignore_index=True)
  extras.iloc[-1, extras.columns.get_loc('Total')] = extras.Total.sum()
  extras_ = extras.iloc[-1]['Total']

  deductions = df.loc[(df['LT#'] == 'Deductions:') | (df['LT#'] == 'Total Deductions:')]
  deductions = deductions[['Project', 'Total']]
  deductions = deductions.loc[deductions.Project != 'Enter Text Here']
  deductions.Total = deductions.Total.astype(float)
  append_Dict1 = {'Project':'Total Extras:'}
  deductions = deductions.append(append_Dict1, ignore_index=True)
  deductions.iloc[-1, deductions.columns.get_loc('Total')] = deductions.Total.sum()
  deductions_ = deductions.iloc[-1]['Total']

  total = df.loc[(df['LT#'] == 'Absolute Total:')]
  total = total[['Project', 'Total']]
  total.iloc[-1, total.columns.get_loc('Project')] = 'Absolute Total:'
  absolute_total = (total_ + extras_) - deductions_
  total.iloc[-1, total.columns.get_loc('Total')] = absolute_total

  df = df.loc[(df['LT#'] != 'Extras:') & (df['LT#'] != 'Deductions:') & (df['LT#'] != 'Absolute Total:')
  & (df['LT#'] != 'Total Extras:') & (df['LT#'] != 'Total Deductions:')]
  
  notes = df[['LT#', 'Notes']]
  notes = notes[df['Notes'] != '-']
  #notes --> loc[row_indexer,col_indexer] = value
  notes.loc[:,'comment'] = df['LT#'] + ': ' + df['Notes']
  comment = notes.comment.tolist()

  formatted_comment = []
  for i in comment:
    formatted_comment.append(f'- {i}')

  # Style Table
  df = df.reset_index()
  df = df.drop(columns=['index', 'Notes'])
  #df = df.rename(columns={"index": ""})
  data = [df.columns.to_list()] + df.values.tolist()
  table = Table(data, hAlign='LEFT')
  table.setStyle(TableStyle([               
      ('BACKGROUND', (0, 0), (13, 0), colors.lightslategray),
      ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
      ('VALIGN', (0, 0), (-1, -1), 'TOP'),
      ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
      ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
      ('ALIGN', (0, 0), (-1, -1), 'CENTER')
  ]))

  # 2) Alternate backgroud color
  rowNumb = len(data)
  for i in range(1, rowNumb):
      if i % 2 == 0:
          bc = colors.white
      else:
          bc = colors.lightgrey
      
      ts = TableStyle(
          [('BACKGROUND', (0,i),(-1,i), bc)]
      )
      table.setStyle(ts)

  # 3) Add borders
  ts = TableStyle(
      [
      ('BOX',(0,0),(-1,-1),1.5,colors.black),
      ('LINEBEFORE',(0, 0), (13, 0),1,colors.black),
      ('LINEABOVE',(0, 0), (13, 1),1,colors.black),
      ('LINEABOVE',(0, -1), (13, -1),1,colors.black),
      ]
  )
  table.setStyle(ts)
  #table.alignment = TA_RIGHT

  # Style Table 2
  extras = extras.reset_index()
  extras = extras.drop(columns=['index'])
  extras = extras.rename(columns={"Project": "Description", 'Total':'Amount'})
  data1 = [extras.columns.to_list()] + extras.values.tolist()
  table1 = Table(data1)
  table1.setStyle(TableStyle([
      ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
      ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
      ('LINEABOVE',(0, 0), (13, 1),1,colors.black),
      ('LINEABOVE',(0, -1), (13, -1),1,colors.black),
      ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
  ]))

  # Style Table 3
  deductions = deductions.reset_index()
  deductions = deductions.drop(columns=['index'])
  deductions = deductions.rename(columns={"Project": "Description", 'Total':'Amount'})
  data2 = [deductions.columns.to_list()] + deductions.values.tolist()
  table2 = Table(data2)
  table2.setStyle(TableStyle([
      ('ALIGN', (0, 0), (-1, -1), "LEFT"),
      ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
      ('LINEABOVE',(0, 0), (13, 1),1,colors.black),
      ('LINEABOVE',(0, -1), (13, -1),1,colors.black),
      ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
  ]))

  # Style Table 4
  total = total.reset_index()
  total = total.drop(columns=['index'])
  data3 = total.values.tolist()
  table3 = Table(data3)
  table3.setStyle(TableStyle([
      ('ALIGN', (0, 0), (-1, -1), "LEFT"),
      ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
      ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
  ]))

  #name1 = FirstName + " " + LastName
  name1 = name.replace("_", " ")
  name1 = name1.replace("Sheet", " ")
  story0 = [Paragraph("Work Sheet", getSampleStyleSheet()['Heading1']),
            Paragraph(f"Employee: {name1}", getSampleStyleSheet()['Heading5']),
            Paragraph(f"Week of: {date}", getSampleStyleSheet()['Heading5']), 
            Spacer(1, 20), table, Spacer(1, 10)]
  
  if len(df) > 10:
    ## fix me
    story0.append(PageBreak())

  for i in range(0,len(notes)):
    if i == 0:
      story0.append(Paragraph(f"Notes:", getSampleStyleSheet()['Heading4']))
    story0.append(Paragraph(f"        {formatted_comment[i]}"))
  
  story1 = [Spacer(1, 10), Paragraph("Extras:", getSampleStyleSheet()['Heading4']),
            table1, Spacer(1, 10),Paragraph("Deductions:", getSampleStyleSheet()['Heading4']),
            table2, Spacer(1, 10), table3,Spacer(1, 40),
            Paragraph("Cheque:_____________", getSampleStyleSheet()['Heading5']), Spacer(1, 10),
            Paragraph("Service Rendered:__________________", getSampleStyleSheet()['Heading5'])]
  story = story0 + story1

  # Use a Frame to dynamically align the compents and write the PDF file
  c = Canvas(f'/content/drive/My Drive/{name}_Report.pdf')
  f = Frame(inch, inch, 6 * inch, 9.75 * inch)
  c.setFont("Helvetica", 10)
  c.drawString(20,20,"Created by: name", )
  f.addFromList(story, c)
  c.save()
  print('Completed.')

# Interaction with user
print('Enter week end date:')
Week_of = input()
# Call function based on inputs
makeReport(Week_of)

