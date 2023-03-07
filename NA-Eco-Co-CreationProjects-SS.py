import json
#import csv
#import xlsxwriter
import os
import smartsheet
import logging

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

def copyProjectsTempalte(newFolderName):
    os.environ['SMARTSHEET_ACCESS_TOKEN'] = 'XXXXXXXXXXXXXXXX'

    templateFolderID=6795484187649924 #Co-creation projects folder id
    destinationFolerID=2755191272433540 #Solutions projects folder id
    taskSheetName='Task Sheet'

    print("Copying template folder ...")
    smart = smartsheet.Smartsheet()
    smart.errors_as_exceptions(True)
    logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

    response = smart.Folders.copy_folder(
    templateFolderID,
    smartsheet.models.ContainerDestination({
        'destination_id': destinationFolerID,
        'destination_type': 'folder',
        'new_name': newFolderName
    }),
    #include='rules,all')
    include='attachments,cell_links,data,discussions,filters,forms,rules,rule_recepients,shares,all')
    print("Copying template folder completed")

    for sheet in smart.Folders.get_folder(response.data.id).sheets:
        if sheet.name == taskSheetName:
            taskSheetId = sheet.id
            break

    taskSheet = smart.Sheets.get_sheet(taskSheetId)

    smart.Sheets.delete_rows(taskSheetId, [row.id for row in taskSheet.rows]) 

    return taskSheetId


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
        
        row.cells.append({
            'column_id': colMap[list(cell.keys())[0]], 
            'value': list(cell.values())[0] if list(cell.values())[0] != None else '',
            'strict': False
            })

    response = smart.Sheets.add_rows(taskSheetId,[row])
    if response.message == 'SUCCESS':
        return response.result[0].id
    else:
        return -1

def mapData(rawData):
    rowData =[]
    rowData.append({"Milestone": rawData['name']})
    rowData.append({"Description": rawData['name']})
    rowData.append({"Predecessors": 2})
    rowData.append({"% Complete": None})
    rowData.append({"Start Date": rawData['start_on']})
    rowData.append({"End Date": rawData['due_on']})
    rowData.append({"Status": ''})
    rowData.append({"Completed": rawData['completed']})
    rowData.append({"Completed At": rawData['completed_at']})
    rowData.append({"Size": ''})
    rowData.append({"Priority": get_custom_field_value("Priority", "e")})
    rowData.append({"Assigned To": (rawData['assignee']['name'] if rawData['assignee'] != None else '')})
    rowData.append({"Duration": ''})
    rowData.append({"Health": ''})
    return rowData

#import_file = "./AsanaData/in/NA-Eco-Co-creation.json"
#import_file = "NA Eco Technical Solution-Co-Creation.json"
import_file = ["./AsanaData/in/NA-Eco-Co-creation.json","./AsanaData/in/NA-Eco-Technical-Solution-Co-Creation.json"]
n=0
for j in range(2):
    with open(import_file[j], "r+", encoding="utf-8") as jsondata:
        data = json.load(jsondata)["data"]
        jsondata.close()

    for i in range(len(data)):
    #for i in range(1):
        x = data[i]
        n+=1
        print("Processing project#"+ str(n) + ": " + x['name'])
        if x['name']== None or len(x['name']) == 0:
             print("Project name missing. Skiping #"+ str(n) + "..." )
             continue
        taskSheetId = copyProjectsTempalte(x['name'][:50])

        rowData =[]
        rowData.append({"Milestone":"Co-Creation"})
        rowData.append({"Description":x['name']})
        rowData.append({"Start Date":x['start_on']})
        rowData.append({"End Date":x['due_on']})
        rowData.append({"Priority":get_custom_field_value("Priority", "e")})
        rowData.append({"Assigned To":(x['assignee']['name'] if x['assignee'] != None else '')})

        parentRowId = writeRows(taskSheetId, rowData, None)
        row_num=2
        
        for j in range(len(x['subtasks'])):
            y=x['subtasks'][j]

            row_num += 1

            parentRowId = writeRows(taskSheetId, mapData(y), None)
            
            for k in range(len(y['subtasks'])):
                z=y['subtasks'][k]
                writeRows(taskSheetId, mapData(z), parentRowId)
print("Done")
