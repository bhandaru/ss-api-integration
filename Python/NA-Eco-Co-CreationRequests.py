import json
import csv
import xlsxwriter
import os

#import_file = "./AsanaData/in/NA-Eco-Co-creation.json"
#export_file = "./AsanaData/out/NA-Eco-Co-creation-Projects.xlsx"
#import_file = "./AsanaData/in/NA-Eco-Technical-Solution-Co-Creation.json"
#export_file = "./AsanaData/out/NA-Eco-Technical-Solution-Co-Creation.xlsx"

import_file = ["./AsanaData/in/NA-Eco-Co-creation.json","./AsanaData/in/NA-Eco-Technical-Solution-Co-Creation.json"]
export_file = "./AsanaData/out/NA-Eco-Technical-Solution-Co-Creation-Combined.xlsx"

with open(export_file, 'w+') as f:
    f.write('{"data": []}')
    f.close()
"""    
with open(import_file, "r+", encoding="utf-8") as jsondata:
    data = json.load(jsondata)["data"]
    jsondata.close()
"""
'''with open("export_file" + ".csv", "r+") as csv_file:
    csvwriter = csv.writer(csv_file, "excel", lineterminator='\n')
    csvwriter.writerow(['name', 'completed', 'assignee'])
    csvwriter.writerows([[x['name'], x['completed'], (x['assignee']['name'] if x['assignee'] != None else '')] for x in data])'''

if os.path.exists(export_file + ".xlsx"):
  os.remove(export_file + ".xlsx")


workbook = xlsxwriter.Workbook(export_file)
 
worksheet = workbook.add_worksheet()
worksheet.write("A1", "Project Name")
worksheet.write("B1", "Description")

worksheet.write("C1", "Leading Partner Type")
worksheet.write("D1", "Salesforce Opp ID(s)")
worksheet.write("E1", "Total Hours")
worksheet.write("F1", "Manager")
worksheet.write("G1", "Co-Creation Progress")
worksheet.write("H1", "Requested By")

worksheet.write("I1", "Primary Ecosystem Solution Partners")
worksheet.write("J1", "Customer Name(s)")
worksheet.write("K1", "Delivery Partner")
worksheet.write("L1", "Delivery Partner - Other")
worksheet.write("M1", "Primary Industry")
worksheet.write("N1", "Technical Contact - Priamry Partner")
worksheet.write("O1", "Technical Contact - Priamry RH Eco Contacts")
worksheet.write("P1", "Executive Summary")
worksheet.write("Q1", "Ideal Customer challenge/problems")
worksheet.write("R1", "Key Customer Benefits")
worksheet.write("S1", "Logical Diagram")
worksheet.write("T1", "Technical Diagram/Reference Arch")
worksheet.write("U1", "Solution Specification/Document")
worksheet.write("V1", "Pre-CRI Solution Write-up")
worksheet.write("W1", "New or Existing")
worksheet.write("X1", "Describe sales process and if it will be partner led or co-sell.")

worksheet.write("Y1", "Approval Status")
worksheet.write("Z1", "Project Status")
worksheet.write("AA1", "Owner")
worksheet.write("AB1", "Priority")
worksheet.write("AC1", "Health")
worksheet.write("AD1", "Start Date")
worksheet.write("AE1", "End Date")
worksheet.write("AF1", "Due Date")
worksheet.write("AG1", "Planned Budget")
worksheet.write("AH1", "Actual Budget")

worksheet.write("AI1", "Project Plan")
worksheet.write("AJ1", "Notes")
worksheet.write("AK1", "Section")
worksheet.write("AL1", "Section")



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

def get_notes_value(field1, field2):
    return d[field1] if field1 in d else (d[field2] if field2 in d else '')

row_num=1
for j in range(2):
    with open(import_file[j], "r+", encoding="utf-8") as jsondata:
        data = json.load(jsondata)["data"]
        jsondata.close()
    
    for i in range(len(data)):
        x = data[i]
        print(x['notes'])
        print(x['name'])
        d=dict(a.split(":\n") for a in x['notes'].split("\n\n")[:-1])

        #d=dict( zip([a.split(":\n")[0]],[ a.split(":\n")[1] if len(a.split(":\n"))>1  else ''])for a in x['notes'].split("\n\n")[:-1])
        
        row_num += 1

        worksheet.write("A" + str(row_num),x['name'])
        worksheet.write("B" + str(row_num),x['name'])

        worksheet.write("C" + str(row_num), get_custom_field_value('Leading Partner Type','e'))
        worksheet.write("D" + str(row_num), get_custom_field_value('SFID(s)','t'))
        worksheet.write("E" + str(row_num), get_custom_field_value('Total Hours','n'))
        worksheet.write("F" + str(row_num), get_custom_field_value('Manager','t'))
        worksheet.write("G" + str(row_num), get_custom_field_value('Co-Creation Progress','e'))
        worksheet.write("H" + str(row_num), get_custom_field_value('Requestor','t'))

        worksheet.write("I" + str(row_num), get_notes_value('Primary Ecosystem Solution Partners','Who is the Primary Partner?')) 
        worksheet.write("J" + str(row_num), get_notes_value('Customer Name(s)', None))
        worksheet.write("K" + str(row_num), get_notes_value('Deilvery Partner', None))
        worksheet.write("L" + str(row_num), get_notes_value('Deilvery Partner - Other','List Other Partners Involved in Solution (if known):'))
        worksheet.write("M" + str(row_num), get_notes_value('Primary Industry','Industry or vertical targeted?'))
        worksheet.write("N" + str(row_num), get_notes_value('Technical Contact - Primary Partner','Who is Partner Lead or Champion?'))
        worksheet.write("O" + str(row_num), get_notes_value('Technical Contact - Primary RH Eco Contacts', None))
        worksheet.write("P" + str(row_num), get_notes_value('Executive Summary','Executive Summary'))
        worksheet.write("Q" + str(row_num), get_notes_value('Ideal Customer challenge/problems', None))
        worksheet.write("R" + str(row_num), get_notes_value('Key Customer Benefits', None))
        worksheet.write("S" + str(row_num), get_notes_value('Logical Diagram', None))
        worksheet.write("T" + str(row_num), get_notes_value('Technical Diagram/Reference Arch', None))
        worksheet.write("U" + str(row_num), get_notes_value('Solution Specification/Document', None))
        worksheet.write("V" + str(row_num), get_notes_value('Pre-CRI Solution Write-up', None))
        worksheet.write("W" + str(row_num), get_notes_value('New or Existing', 'Is this a new solution or this an existing solution?'))
        worksheet.write("X" + str(row_num), get_notes_value('Describe sales process and if it will be partner led or co-sell.', None))

        worksheet.write("Y" + str(row_num), '')
        worksheet.write("Z" + str(row_num), get_custom_field_value('Co-Creation Progress','e'))
        worksheet.write("AA" + str(row_num),(x['assignee']['name'] if x['assignee'] != None else ''))
        worksheet.write("AB" + str(row_num), get_custom_field_value('Priority','e'))
        worksheet.write("AC" + str(row_num), '')
        worksheet.write("AD" + str(row_num),x['start_on'])
        worksheet.write("AE" + str(row_num),x['due_on'])
        worksheet.write("AF" + str(row_num),x['due_on'])
        worksheet.write("AG" + str(row_num),0)
        worksheet.write("AH" + str(row_num),0)
        worksheet.write("AI" + str(row_num),'')
        worksheet.write("AJ" + str(row_num),x['notes'])
        worksheet.write("AK" + str(row_num),x['memberships'][0]['section']['name'])
        worksheet.write("AL" + str(row_num),x['created_at'])




workbook.close()