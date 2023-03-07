import json
import csv
import xlsxwriter
import os
import smartsheet
import logging
from NAEcoCoCreationProjectsSSclass import mySMProject

#import_file = "./AsanaData/in/NA-Eco-Co-creation.json"
#export_file = "./AsanaData/out/NA-Eco-Co-creation-Projects.xlsx"
#import_file = "./AsanaData/in/NA-Eco-Technical-Solution-Co-Creation.json"
#export_file = "./AsanaData/out/NA-Eco-Technical-Solution-Co-Creation.xlsx"

import_file = ["./AsanaData/in/NA-Eco-Co-creation.json","./AsanaData/in/NA-Eco-Technical-Solution-Co-Creation.json"]
export_file = "./AsanaData/out/NA-Eco-Technical-Solution-Co-Creation-Combined.xlsx"

projectsSheetID=1949511405856644 #Co-creation Projects sheet id
os.environ['SMARTSHEET_ACCESS_TOKEN'] = 'XXXXXXXXXXXXXXXX'

#Connect to Smartsheet
smart = smartsheet.Smartsheet()
smart.errors_as_exceptions(True)
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

#Get Project Sheet and delete rows
taskSheet = smart.Sheets.get_sheet(projectsSheetID)
if len(taskSheet.rows) > 0:
    smart.Sheets.delete_rows(projectsSheetID, [row.id for row in taskSheet.rows]) 

with open(export_file, 'w+') as f:
    f.write('{"data": []}')
    f.close()
"""    
with open(import_file, "r+", encoding="utf-8") as jsondata:
    data = json.load(jsondata)["data"]
    jsondata.close()
"""
'''
with open("export_file" + ".csv", "r+") as csv_file:
    csvwriter = csv.writer(csv_file, "excel", lineterminator='\n')
    csvwriter.writerow(['name', 'completed', 'assignee'])
    csvwriter.writerows([[x['name'], x['completed'], (x['assignee']['name'] if x['assignee'] != None else '')] for x in data])
'''


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


def get_custom_field_value(field, val_type, f):
    filter_val = list(filter(lambda dictionary: dictionary['name'] == field, f['custom_fields']))
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

def get_notes_value(field1, field2, dNotes):
    return dNotes[field1] if field1 in dNotes else (dNotes[field2] if field2 in dNotes else '')

def mapData(rawData, projectURL):

    d=dict(a.split(":\n") for a in rawData['notes'].split("\n\n")[:-1])

    rowData =[]
    rowData.append({"Project Name":rawData['name']})
    rowData.append({"Description":rawData['name']})

    rowData.append({"Leading Partner Type":get_custom_field_value('Leading Partner Type','e',rawData)})
    rowData.append({"Salesforce Opp ID(s)":get_custom_field_value('SFID(s)','t',rawData)})
    rowData.append({"Total Hours":get_custom_field_value('Total Hours','n',rawData)})
    rowData.append({"Manager":get_custom_field_value('Manager','t',rawData)})
    rowData.append({"Co-Creation Progress":get_custom_field_value('Co-Creation Progress','e',rawData)})
    rowData.append({"Requested By":get_custom_field_value('Requestor','t',rawData)})

    rowData.append({"Primary Ecosystem Solution Partners":get_notes_value('Primary Ecosystem Solution Partners','Who is the Primary Partner?',d)})
    rowData.append({"Customer Name(s)":get_notes_value('Customer Name(s)', None,d)})
    rowData.append({"Delivery Partner":get_notes_value('Deilvery Partner', None,d)})
    rowData.append({"Delivery Partner - Other":get_notes_value('Deilvery Partner - Other','List Other Partners Involved in Solution (if known):',d)})
    rowData.append({"Primary Industry":get_notes_value('Primary Industry','Industry or vertical targeted?',d)})
    rowData.append({"Technical Contact - Primary Partner":get_notes_value('Technical Contact - Primary Partner','Who is Partner Lead or Champion?',d)})
    rowData.append({"Technical Contact - Primary RH Eco Contacts":get_notes_value('Technical Contact - Primary RH Eco Contacts', None,d)})
    rowData.append({"Executive Summary":get_notes_value('Executive Summary','Executive Summary',d)})
    rowData.append({"Ideal Customer challenge/problems":get_notes_value('Ideal Customer challenge/problems', None,d)})
    rowData.append({"Key Customer Benefits":get_notes_value('Key Customer Benefits', None,d)})
    rowData.append({"Logical Diagram":get_notes_value('Logical Diagram', None,d)})
    rowData.append({"Technical Diagram/Reference Arch":get_notes_value('Technical Diagram/Reference Arch', None,d)})
    rowData.append({"Solution Specification/Document":get_notes_value('Solution Specification/Document', None,d)})
    rowData.append({"Pre-CRI Solution Write-up":get_notes_value('Pre-CRI Solution Write-up', None,d)})
    rowData.append({"New or Existing":get_notes_value('New or Existing', 'Is this a new solution or this an existing solution?',d)})
    rowData.append({"Describe sales process and if it will be partner led or co-sell.":get_notes_value('Describe sales process and if it will be partner led or co-sell.', None,d)})

    rowData.append({"Approval Status":''})
    rowData.append({"Project Status":get_custom_field_value('Co-Creation Progress','e',rawData)})
    rowData.append({"Owner":(rawData['assignee']['name'] if rawData['assignee'] != None else '')})
    rowData.append({"Priority":get_custom_field_value('Priority','e',rawData)})
    rowData.append({"Health":''})
    rowData.append({"Create Date":rawData['created_at'][:10] if rawData['created_at'] != None else ''})
    rowData.append({"Start Date":rawData['start_on']})
    rowData.append({"End Date":rawData['due_on']})
    rowData.append({"Due Date":rawData['due_on']})
    #rowData.append({"Planned Budget":0})
    #rowData.append({"Actual Budget":0})
    rowData.append({"Completed":rawData['completed']})
    rowData.append({"Completed Date":rawData['completed_at'][:10] if rawData['completed_at']!= None else ''})


    rowData.append({"Project Plan":projectURL})
    rowData.append({"Notes":rawData['notes']})
    rowData.append({"Section":rawData['memberships'][0]['section']['name']})
    
    return rowData              

def writeRows(taskSheetId, rowData, parentId):
    #Get taskSheet
    smart = smartsheet.Smartsheet()
    taskSheet = smart.Sheets.get_sheet(taskSheetId)
    
    #Get Column IDs
    colMap = dict((col.title,col.id) for col in taskSheet.columns)
   
    #Write row
    rows = []
    row = smart.models.Row()
    row.to_bottom = True
    if parentId != None:
        row.parent_id = parentId 

    for cell in rowData:
        
        if(list(cell.keys())[0] == "Project Plan"):
            row.cells.append({
            'column_id': colMap[list(cell.keys())[0][:50]], 
            'value'  : 'Project Dashboard',
            'hyperlink' : { "url": list(cell.values())[0] if list(cell.values())[0] != None else ''},
            'strict': False
            })
        else:
            row.cells.append({
                'column_id': colMap[list(cell.keys())[0][:50]], 
                'value': list(cell.values())[0] if list(cell.values())[0] != None else '',
                'strict': False
                })

    response = smart.Sheets.add_rows(taskSheetId,[row])
    if response.message == 'SUCCESS':
        return response.result[0].id
    else:
        return -1

row_num=1
for j in range(2):
    with open(import_file[j], "r+", encoding="utf-8") as jsondata:
        data = json.load(jsondata)["data"]
        jsondata.close()
    
    for i in range(len(data)):
        x = data[i]
        print(x['notes'])
        d=dict(a.split(":\n") for a in x['notes'].split("\n\n")[:-1])

        #d=dict( zip([a.split(":\n")[0]],[ a.split(":\n")[1] if len(a.split(":\n"))>1  else ''])for a in x['notes'].split("\n\n")[:-1])
        
        cp = mySMProject()
        dashboardlink = cp.go(x)


        row_num += 1

        worksheet.write("A" + str(row_num),x['name'])
        worksheet.write("B" + str(row_num),x['name'])

        worksheet.write("C" + str(row_num), get_custom_field_value('Leading Partner Type','e',x))
        worksheet.write("D" + str(row_num), get_custom_field_value('SFID(s)','t',x))
        worksheet.write("E" + str(row_num), get_custom_field_value('Total Hours','n',x))
        worksheet.write("F" + str(row_num), get_custom_field_value('Manager','t',x))
        worksheet.write("G" + str(row_num), get_custom_field_value('Co-Creation Progress','e',x))
        worksheet.write("H" + str(row_num), get_custom_field_value('Requestor','t',x))

        worksheet.write("I" + str(row_num), get_notes_value('Primary Ecosystem Solution Partners','Who is the Primary Partner?',d)) 
        worksheet.write("J" + str(row_num), get_notes_value('Customer Name(s)', None,d))
        worksheet.write("K" + str(row_num), get_notes_value('Deilvery Partner', None,d))
        worksheet.write("L" + str(row_num), get_notes_value('Deilvery Partner - Other','List Other Partners Involved in Solution (if known):',d))
        worksheet.write("M" + str(row_num), get_notes_value('Primary Industry','Industry or vertical targeted?',d))
        worksheet.write("N" + str(row_num), get_notes_value('Technical Contact - Primary Partner','Who is Partner Lead or Champion?',d))
        worksheet.write("O" + str(row_num), get_notes_value('Technical Contact - Primary RH Eco Contacts', None,d))
        worksheet.write("P" + str(row_num), get_notes_value('Executive Summary','Executive Summary',d))
        worksheet.write("Q" + str(row_num), get_notes_value('Ideal Customer challenge/problems', None,d))
        worksheet.write("R" + str(row_num), get_notes_value('Key Customer Benefits', None,d))
        worksheet.write("S" + str(row_num), get_notes_value('Logical Diagram', None,d))
        worksheet.write("T" + str(row_num), get_notes_value('Technical Diagram/Reference Arch', None,d))
        worksheet.write("U" + str(row_num), get_notes_value('Solution Specification/Document', None,d))
        worksheet.write("V" + str(row_num), get_notes_value('Pre-CRI Solution Write-up', None,d))
        worksheet.write("W" + str(row_num), get_notes_value('New or Existing', 'Is this a new solution or this an existing solution?',d))
        worksheet.write("X" + str(row_num), get_notes_value('Describe sales process and if it will be partner led or co-sell.', None,d))

        worksheet.write("Y" + str(row_num), '')
        worksheet.write("Z" + str(row_num), get_custom_field_value('Co-Creation Progress','e',x))
        worksheet.write("AA" + str(row_num),(x['assignee']['name'] if x['assignee'] != None else ''))
        worksheet.write("AB" + str(row_num), get_custom_field_value('Priority','e',x))
        worksheet.write("AC" + str(row_num), '')
        worksheet.write("AD" + str(row_num),x['start_on'])
        worksheet.write("AE" + str(row_num),x['due_on'])
        worksheet.write("AF" + str(row_num),x['due_on'])
        worksheet.write("AG" + str(row_num),0)
        worksheet.write("AH" + str(row_num),0)
        worksheet.write("AI" + str(row_num),'')
        worksheet.write("AJ" + str(row_num),x['notes'])
        worksheet.write("AK" + str(row_num),x['memberships'][0]['section']['name'])

        rowID = writeRows(projectsSheetID, mapData(x,dashboardlink), None)

        #writeRows(taskSheetId, rowData, None)


workbook.close()
