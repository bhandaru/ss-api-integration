import json
import csv
import xlsxwriter
import os

import_file = "Asana-Co-creation.json"
export_file = "Eco Co-creation Projects"
with open(export_file, 'w+') as f:
    f.write('{"data": []}')
    f.close()
with open(import_file, "r+", encoding="utf-8") as jsondata:
    data = json.load(jsondata)["data"]
    jsondata.close()
'''with open("export_file" + ".csv", "r+") as csv_file:
    csvwriter = csv.writer(csv_file, "excel", lineterminator='\n')
    csvwriter.writerow(['name', 'completed', 'assignee'])
    csvwriter.writerows([[x['name'], x['completed'], (x['assignee']['name'] if x['assignee'] != None else '')] for x in data])'''

if os.path.exists(export_file + ".xlsx"):
  os.remove(export_file + ".xlsx")


workbook = xlsxwriter.Workbook(export_file + ".xlsx")
 
worksheet = workbook.add_worksheet()
worksheet.write("A1", "Project Name")
worksheet.write("B1", "Description")
worksheet.write("C1", "Leading Partner Type")
worksheet.write("D1", "Salesforce Opp ID(s)")
worksheet.write("E1", "Total Hours")
worksheet.write("F1", "Manager")
worksheet.write("G1", "Co-Creation Progress")
worksheet.write("H1", "Requestor")
worksheet.write("I1", "Approval Status")
worksheet.write("J1", "Project Status")
worksheet.write("K1", "Owner")
worksheet.write("L1", "Priority")
worksheet.write("M1", "Health")
worksheet.write("N1", "Start Date")
worksheet.write("O1", "End Date")
worksheet.write("P1", "Due Date")
worksheet.write("Q1", "Project Plan")
worksheet.write("R1", "Notes")
worksheet.write("S1", "Section")

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
        case _:
            return filter_val['text_value']
    return 


for i in range(len(data)):
    x = data[i]

    worksheet.write("A" + str(i+2),x['name'])
    worksheet.write("B" + str(i+2),'')
    worksheet.write("C" + str(i+2), get_custom_field_value('Leading Partner Type','e'))
    worksheet.write("D" + str(i+2), get_custom_field_value('SFID(s)','t'))
    worksheet.write("E" + str(i+2), get_custom_field_value('Total Hours','n'))
    worksheet.write("F" + str(i+2), get_custom_field_value('Manager','t'))
    worksheet.write("G" + str(i+2), get_custom_field_value('Co-Creation Progress','e'))
    worksheet.write("H" + str(i+2), get_custom_field_value('Requestor','t'))
    worksheet.write("I" + str(i+2), '')
    worksheet.write("J" + str(i+2), '')
    worksheet.write("K" + str(i+2),(x['assignee']['name'] if x['assignee'] != None else ''))
    worksheet.write("L" + str(i+2), get_custom_field_value('Priority','e'))
    worksheet.write("M" + str(i+2), get_custom_field_value('Co-Creation Progress','e'))
    worksheet.write("N" + str(i+2),x['start_on'])
    worksheet.write("O" + str(i+2),x['due_on'])
    worksheet.write("P" + str(i+2),x['due_on'])
    worksheet.write("Q" + str(i+2),'')
    worksheet.write("R" + str(i+2),x['notes'])
    worksheet.write("S" + str(i+2),x['memberships'][0]['section']['name'])

    



workbook.close()