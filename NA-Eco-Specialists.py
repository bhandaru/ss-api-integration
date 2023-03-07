import json
import csv
import xlsxwriter
import os

import_file = "Eco-Specialist-SA.json"
#import_file = "NA Eco Technical Solution-Co-Creation.json"


with open("Data/In/"+import_file, "r+", encoding="utf-8") as jsondata:
    data = json.load(jsondata)["data"]
    jsondata.close()
'''with open("export_file" + ".csv", "r+") as csv_file:
    csvwriter = csv.writer(csv_file, "excel", lineterminator='\n')
    csvwriter.writerow(['name', 'completed', 'assignee'])
    csvwriter.writerows([[x['name'], x['completed'], (x['assignee']['name'] if x['assignee'] != None else '')] for x in data])'''



def get_custom_field_value(field, val_type):
    filter_val = list(filter(lambda dictionary: dictionary['name'] == field, x['custom_fields']))
    if filter_val == []:
        return ''
    filter_val = filter_val[0]
    match val_type:
        case 'e':
            return (filter_val['enum_value']['name'] if filter_val['enum_value'] != None else '')
        case 'n':
            return filter_val['number_value']
        case 'me':
            return '~'.join([item['name'] for item in filter_val['multi_enum_values']])
        case _:
            return filter_val['text_value']
    return 

#export_file = "Data/Out/"+x['name'].replace('/','-') + ".xlsx" #"Eco Co-creation Projects"
export_file = "Data/Out/Eco-Specialist-SA.xlsx"

if os.path.exists(export_file):
    os.remove(export_file)

with open(export_file, 'w+') as f:
    f.write('{"data": []}')
    f.close()

workbook = xlsxwriter.Workbook(export_file)
"""
cf = workbook.add_format({'bold': True,'bg_color': 'navy','font_color': 'white'})
cfdate = workbook.add_format({'num_format': "mm/dd/yyyy", 'bold': True,'bg_color': 'navy','font_color': 'white'})
cfnum = workbook.add_format({'num_format': '#.##%','bold': True,'bg_color': 'navy','font_color': 'white'})
datefmt = workbook.add_format({'num_format': "mm/dd/yyyy"})
"""
cf = workbook.add_format(None)
cfdate = workbook.add_format(None)
cfnum = workbook.add_format(None)
datefmt = workbook.add_format({'num_format': "mm/dd/yyyy"})

worksheet = workbook.add_worksheet()
worksheet.write("A1", "Request Title")
worksheet.write("B1", "Assignee")
worksheet.write("C1", "Due Date")
worksheet.write("D1", "Completed")
worksheet.write("E1", "Completed At")
worksheet.write("F1", "Created At")
worksheet.write("G1", "Section")
worksheet.write("H1", "Requestor")
worksheet.write("I1", "Followers")
worksheet.write("J1", "Task Progress")
worksheet.write("K1", "Account Number / Effort Name")
worksheet.write("L1", "Profile Requested")
worksheet.write("M1", "Primary Technology Requested")
worksheet.write("N1", "Type of work (SSA)")
worksheet.write("O1", "On Site / Remote")
worksheet.write("P1", "Priority Level?")
worksheet.write("Q1", "Notes")

row_num=1

for i in range(len(data)):
#for i in range(1):
    x = data[i]
    row_num += 1
    worksheet.write("A"+str(row_num), x['name'])
    worksheet.write("B"+str(row_num), (x['assignee']['name'] if x['assignee'] != None else ''))
    worksheet.write("C"+str(row_num), x['due_on'])
    worksheet.write("D"+str(row_num), x['completed'])
    worksheet.write("E"+str(row_num), x['completed_at'])
    worksheet.write("E"+str(row_num), x['created_at'])
    worksheet.write("F"+str(row_num), x['due_on'],cfdate)
    worksheet.write("G"+str(row_num), x['memberships'][0]['section']['name'])
    worksheet.write("H"+str(row_num), get_custom_field_value('Requestor','t'))
    worksheet.write("I"+str(row_num), ','.join([item['name'] for item in x['followers']]))
    worksheet.write("J"+str(row_num), get_custom_field_value('Task Progress','e'))
    worksheet.write("K"+str(row_num), get_custom_field_value('Account Number / Effort Name','t'))
    worksheet.write("L"+str(row_num), get_custom_field_value('Profile Requested','e'))
    worksheet.write("M"+str(row_num), get_custom_field_value('Primary Technology Requested','e'))
    worksheet.write("N"+str(row_num), get_custom_field_value('Type of work (SSA)','me'))
    worksheet.write("O"+str(row_num), get_custom_field_value('On Site / Remote','e'))
    worksheet.write("P"+str(row_num), get_custom_field_value('Priority level?','e'))
    worksheet.write("Q"+str(row_num), x['notes'])

    print(str(row_num) + " - " + x['name'])
  

    if len(x['subtasks']) > 0:
        worksheet.set_row(row_num-1, None, None, {'level': 1})
        print(str(row_num) + ", "+ str(1))
    else:
        worksheet.set_row(row_num-1, None, None, {'level': 0})
        print(str(row_num) + ", "+ str(0))
    worksheet.set_row(row_num-1, None, None, {'level': 0})
    

    
    #row_groups = [[2], [3, 7, 10, 15, 26, 33]]
    
    for j in range(len(x['subtasks'])):
        y=x['subtasks'][j]
        #cell_indent_format = workbook.add_format()
        #cell_indent_format.set_indent(2)
        
        row_num += 1

        #worksheet.set_row(row_num-1, None, cell_format=cf)

        worksheet.write("A"+str(row_num), y['name'])
        worksheet.write("B"+str(row_num), (y['assignee']['name'] if y['assignee'] != None else ''))
        worksheet.write("C"+str(row_num), y['due_on'])
        worksheet.write("D"+str(row_num), y['completed'])
        worksheet.write("E"+str(row_num), y['completed_at'])
        worksheet.write("E"+str(row_num), y['created_at'])
        worksheet.write("F"+str(row_num), y['due_on'],cfdate)
        worksheet.write("G"+str(row_num), y['memberships'][0]['section']['name'] if len(y['memberships']) > 0 else '')
        worksheet.write("H"+str(row_num), get_custom_field_value('Requestor','t'))
        worksheet.write("I"+str(row_num), ','.join([item['name'] for item in y['followers']]))
        worksheet.write("J"+str(row_num), get_custom_field_value('Task Progress','e'))
        worksheet.write("K"+str(row_num), get_custom_field_value('Account Number / Effort Name','t'))
        worksheet.write("L"+str(row_num), get_custom_field_value('Profile Requested','e'))
        worksheet.write("M"+str(row_num), get_custom_field_value('Primary Technology Requested','e'))
        worksheet.write("N"+str(row_num), get_custom_field_value('Type of work (SSA)','me'))
        worksheet.write("O"+str(row_num), get_custom_field_value('On Site / Remote','e'))
        worksheet.write("P"+str(row_num), get_custom_field_value('Priority level?','e'))
        worksheet.write("Q"+str(row_num), y['notes'])

        print(str(row_num) + ", "+ str(3-1))

        worksheet.set_row((row_num-1), None, None, {'level': 1})

workbook.close()