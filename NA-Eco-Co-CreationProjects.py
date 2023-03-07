import json
import csv
import xlsxwriter
import os
import smartsheet
import logging

import_file = "./AsanaData/in/NA-Eco-Co-creation.json"
#import_file = "NA Eco Technical Solution-Co-Creation.json"

with open(import_file, "r+", encoding="utf-8") as jsondata:
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
        case _:
            return filter_val['text_value']
    return 

def copyProjectsTempalte(newFolderName):
    os.environ['SMARTSHEET_ACCESS_TOKEN'] = 'jvmfs8OA00K7iJMD2tapJAl7U4jk6tUHqs3Lh'

    templateFolderID=8346295946504068
    destinationFolerID=2755191272433540

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
    print("Done")


for i in range(len(data)):
#for i in range(1):
    x = data[i]

    export_file = "./AsanaData/out/"+x['name'].replace('/','-') + ".xlsx" #"Eco Co-creation Projects"

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
    worksheet.write("A1", "Milestone")
    worksheet.write("B1", "Description")
    worksheet.write("C1", "Predecessors")
    worksheet.write("D1", "% Complete")
    worksheet.write("E1", "Start Date")
    worksheet.write("F1", "End Date")
    worksheet.write("G1", "Status")
    worksheet.write("H1", "Completed")
    worksheet.write("I1", "Completed At")
    worksheet.write("J1", "Size")
    worksheet.write("K1", "Priority")
    worksheet.write("L1", "Assigned To")
    worksheet.write("M1", "Duration")
    worksheet.write("N1", "Health")

    worksheet.set_row(2-1, None, cell_format=cf)
    worksheet.set_column('E:F', None, datefmt)
    worksheet.write("A"+str(2), "Co-Creation")
    worksheet.write("B"+str(2), x['name'])
    worksheet.write("E"+str(2), x['start_on'],cfdate)
    worksheet.write("F"+str(2), x['due_on'],cfdate)
    worksheet.write("K"+str(2), get_custom_field_value("Priority", "e"))
    worksheet.write("L"+str(2), (x['assignee']['name'] if x['assignee'] != None else ''),cfdate)
  

    #worksheet.set_row(2-1, None, None, {'level': 1})
    #print(str(2-1) + ", "+ str(1))

    row_num=2
    #row_groups = [[2], [3, 7, 10, 15, 26, 33]]
    
    for j in range(len(x['subtasks'])):
        y=x['subtasks'][j]
        #cell_indent_format = workbook.add_format()
        #cell_indent_format.set_indent(2)
        
        row_num += 1

        worksheet.set_row(row_num-1, None, cell_format=cf)

        worksheet.write("A"+str(row_num), y['name'],cf)
        worksheet.write("B"+str(row_num), y['name'],cf)
        worksheet.write("C"+str(row_num), 2, cfnum)
        worksheet.write("D"+str(row_num), None, cf)
        worksheet.write("E"+str(row_num), y['start_on'],cfdate)
        worksheet.write("F"+str(row_num), y['due_on'],cfdate)
        worksheet.write("G"+str(row_num), '',cf)
        worksheet.write("H"+str(row_num), y['completed'],cf)
        worksheet.write("I"+str(row_num), y['completed_at'],cfdate)
        worksheet.write("J"+str(row_num), '',cf)
        worksheet.write("K"+str(row_num), get_custom_field_value("Priority", "e"))
        worksheet.write("L"+str(row_num), (y['assignee']['name'] if y['assignee'] != None else ''),cf)
        worksheet.write("M"+str(row_num), '',cf)

        print(str(row_num-1) + ", "+ str(2-1))

        worksheet.set_row((row_num-1), None, None, {'level': 2-1})
        
        for k in range(len(y['subtasks'])):
            z=y['subtasks'][k]
            #cell_indent_format = workbook.add_format()
            #cell_indent_format.set_indent(2)

            date_format = workbook.add_format({'num_format': "%Y-%m-%d"})
            number_format = workbook.add_format({'num_format': '#.##%'})

            #worksheet.write("A"+str(k+j+4), z['name'])
            row_num += 1
            worksheet.set_row(row_num-1, None, None, {'level': 3-1})
            print(str(row_num-1) + ", "+ str(3-1))
            worksheet.write("B"+str(row_num), z['name'])
            worksheet.write("C"+str(row_num), None, number_format)
            worksheet.write("D"+str(row_num), '')
            worksheet.write("E"+str(row_num), z['start_on'],date_format)
            worksheet.write("F"+str(row_num), z['due_on'],date_format)
            worksheet.write("G"+str(row_num), '')
            worksheet.write("H"+str(row_num), z['completed'])
            worksheet.write("I"+str(row_num), z['completed_at'],date_format)
            worksheet.write("J"+str(row_num), '')
            worksheet.write("K"+str(row_num), '')
            worksheet.write("L"+str(row_num), (z['assignee']['name'] if z['assignee'] != None else ''))
            worksheet.write("M"+str(row_num), '')
            worksheet.write("N"+str(row_num), '')

    workbook.close()

    copyProjectsTempalte(x['name'][:50])