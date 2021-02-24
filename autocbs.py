#  Copyright (c) 2021 Freyr Yggdrasil 
#  https://github.com/FreyrYggdrasil/autocbs

#  Permission is hereby granted, free of charge, to any person
#  obtaining a copy of this software and associated documentation
#  files (the "Software"), to deal in the Software without
#  restriction, including without limitation the rights to use,
#  copy, modify, merge, publish, distribute, sublicense, and/or sell
#  copies of the Software, and to permit persons to whom the
#  Software is furnished to do so, subject to the following
#  conditions:

#  The above copyright notice and this permission notice shall be
#  included in all copies or substantial portions of the Software.

#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
#  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
#  OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
#  NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
#  HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
#  WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#  FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
#  OTHER DEALINGS IN THE SOFTWARE.

#  main import modules
import json
import cbsodata

#  **************************************************
#  from typing import List
import pickle
import itertools as it

#  format output file and file name and running time
import datetime 
import csv
import os
import glob

#  dataframes
import pandas as pd
import numpy as np
import functools
import operator
from pathlib import Path

#  excel support
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

#  command line arguments
import sys

#  **************************************************
#  timing program
begin = datetime.datetime.now()

#  **************************************************
#  control information
#  to do:
#  extend with date {lastday}, {lastweek}, {lastmonth}
controlInformationTables = {}
controlInformationTable = {
    "Title":"","Updated":"","ShortTitle":"","Identifier":"",
    "Summary":"","Modified":"","ReasonDelivery":"",
    "Frequency":"","Period":"","RecordCount":"","lastRefreshDate":"",
    "lastRefreshDateJson":"","lastRefreshDateCsv":"",
    "lastRefreshDateExcel":""}
control_file = 'get_data_control.xlsx'

# **************************************************
# get cmnd arguments
# init - want to be able to use just switch
help_switch = False         #0 -h 
download_data = False       #1 -d download data
folder_name = ''            #2 -f ./data/ is default
search_arg = []             #3 -s <keywords,> to search for in Shortdescription
table_identifier = []       #4 -i <id,>
get_tables = 0              #5 -n <nr>
output_level = ''           #6 -v silent,info,warning,error,critical # to do
start_record = 0            #7 -b <nr> begin at record 
table_meta = False          #8 -m only download TableInfo
no_master = False           #9 -nm do not maintain a master xlsx file
table_prop = False          #10 -p only retrieve DataProperties incl. TableInfos
download_csv = False        #11 -csv download data as csv file
data_refresh = False        #12 -update if True update existing data
download_excel = False      #13 -xls download data as xlsx file
download_json =  False      #14 -json  download data as json file
force_download = False      #15 -force force update of local data
modified_within = ''        #16 -md last[day,week,month,year] only tables that are changed

# arguments
argument_list = [['-h', help_switch], ['-d', download_data], 
    ['-f', folder_name], ['-s', search_arg], ['-id', table_identifier], 
    ['-n', get_tables], ['-v', output_level], ['-b', start_record], 
    ['-m', table_meta], ['-nm', no_master], ['-p', table_prop], 
    ['-csv', download_csv], ['-update', data_refresh], 
    ['-xls', download_excel], ['-json', download_json], 
    ['-force', force_download], 
    ['-mw', modified_within]]

# **************************************************
# evaluate command line arguments
argument_list = [list((i, argument_list[i])) for i in range(len(argument_list))]
sys_args = list(sys.argv)
length = len(argument_list)
results = []
for a in range(len(sys_args)):
    for x in range(length):
        try:
            result = argument_list[x][1].index(sys_args[a])            
            if type(argument_list[x][1][1]) == type(str()):
                results.append([x, sys_args[a], sys_args[a+1]])
            elif type(argument_list[x][1][1]) == type(bool()):
                results.append([x, sys_args[a], True])
            elif type(argument_list[x][1][1]) == type(int()):
                results.append([x, sys_args[a], int(sys_args[a+1])])
            elif type(argument_list[x][1][1]) == type(list()):
                results.append([x, sys_args[a], sys_args[a+1].split(',')])
        except ValueError:
            pass
        except IndexError:
            if type(argument_list[x][1][1]) == type(bool()):
                results.append([x, sys_args[a], True])            
            pass

# **************************************************
# assign arguments
for a in range(len(results)): 
    argument_list[int(str(results[a][0]))][1][1] = results[a][2]
    
help_switch = argument_list[0][1][1]
download_data = argument_list[1][1][1]
folder_name = argument_list[2][1][1]
search_arg = argument_list[3][1][1]
table_identifier = argument_list[4][1][1]
get_tables = argument_list[5][1][1]
output_level = argument_list[6][1][1]    
start_record = argument_list[7][1][1]
table_meta = argument_list[8][1][1]
no_master = argument_list[9][1][1]
table_prop = argument_list[10][1][1]
download_csv = argument_list[11][1][1]
data_refresh = argument_list[12][1][1]
download_excel = argument_list[13][1][1]
download_json = argument_list[14][1][1]
force_download = argument_list[15][1][1]
modified_within = argument_list[16][1][1]

if help_switch:
    print(sys.argv[0], ':: Download CBS data tables\narguments:\n-h\t\t\tthis help\n-d\t\t\tdownload data (in -f)\n-f <folder>\t\tfolder name for data (+ Identifier), ./data/is the default\n-s <string,>\t\tsearch for keywords in table shortdescription (can be comma seperated)\n-id <identifier>\ttable to download\n-v <level>\t\toutput level silent, info, warning, error, critical\n-n <nr>\t\t\tmaximum tables to get\n-b <nr>\t\t\tstart at record\n-m\t\t\tGet meta data table\n-nm\t\t\tdo NOT maintain master excel file with table info\n-p\t\t\tget the info of the table\n-csv\t\t\tsave files as csv\n-force\t\t\tForce download of large result set\n-update\t\t\tUpdate already downloaded tables\n-xls\t\t\tupdate excel file\n-json\t\t\tupdate json files\n-mw\t\t\ttable modifed within lastday, lastweek or lastmonth')
    raise SystemExit(0)

# loglevel
# todo
levels = ['silent', 'critical', 'error', 'warning', 'info', 'verbose', 'allmsg']
silent = 0
critical = 1
error = 2
warning = 3
info = 4
verbose = 5
allmsg = 6

for o in range(len(levels)):
    if levels[o] == str(output_level):
        log_level = o
    else:
        log_level = 4   # info

# **************************************************
# print string to screen for user feedback
def p(plevel,text,*args):
    
    no_linefeed=False
    if not text: 
        text = ''
    else:
        text = str(text)
        
    try:
        for i in args:
            if not i == 'end=""':
                text = text + ' ' + str(i)
            else:
                no_linefeed = True

        if log_level >= plevel: 
            if no_linefeed:
                print(text, end="")
            else:
                print(text)
        else:
            pass
    except ValueError:
        pass
    return

# who we are and what we do
p(verbose,str(sys.argv).replace('[','').replace(']','').replace("'",''),'\n')

# general download settings (default is all)
if download_data and (not download_csv and not download_excel and not download_json): 
    download_csv = True 
    download_excel = True 
    download_json = True

# must have -d
if not download_data and (download_csv or download_excel or download_json):
    download_data = True

# only update local master when actually downloading data
if not download_data: no_master = True


# default output folder
if folder_name == '':
    folder_name = './data/'

# **************************
# excel helpers
def transpose(ws, min_row, max_row, min_col, max_col):
    for row in range(min_row, max_row+1):
        for col in range(min_col, max_col+1):
            ws.cell(row=col,column=row).value = ws.cell(row=row,column=col).value

def transpose_row_to_col(ws, min_row, max_row, min_col, max_col, target_cell_address=(1,1), delete_source=False):
    cell_values = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell_values.append(cell.value)
            if delete_source:
                cell.value = ""
    fill_cells(ws, target_cell_address[0], target_cell_address[1], cell_values)

def transpose_col_to_row(ws, min_row, max_row, min_col, max_col, target_cell_address=(1,1), delete_source=False):
    cell_values = []
    for col in ws.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in col:
            cell_values.append(cell.value)
            if delete_source:
                cell.value = ""
    fill_cells(ws, target_cell_address[0], target_cell_address[1], cell_values)

def fill_cells(ws, start_row, start_column, cell_values):
    row = start_row
    column = start_column
    for value in cell_values:
        ws.cell(row=row,column=column).value = value
        row += 1

def convertTuple(tup): 
    str = functools.reduce(operator.add, (tup)) 
    return str

# end excel helpers      
     
# *********************************
# save objects
def save_data(data, dir, p_identifier, metadata_name, argument):

    if type(argument) == type(str()):
        output_file = os.path.join(dir, p_identifier+'-'+metadata_name + '.' + argument)
        
    elif type(data) == type(dict()):
        # updating master file
        # todo
        pass
        
    else:
        # getting data for excel
        output_file = os.path.join(dir, p_identifier+'-objects.xlsx')
        sheet = argument

    if argument == 'json':
        my_data = json.loads(str(data))
        with open(output_file, 'w') as output_file:
            json.dump(my_data, output_file, indent=4)
        output_file.close()
        
        # update date mc
        if not no_master:
            controlInformationTable['lastRefreshDateJson'] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        
        return my_data
        
    elif argument == 'csv':
        data_csv = data.to_csv(output_file, sep=";", na_rep="", quoting=csv.QUOTE_ALL, quotechar='"', doublequote=True, escapechar="\\", index = False)
        
        # update date mc
        if not no_master:
            controlInformationTable['lastRefreshDateCsv'] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

        return data_csv
    
    else:
        # excel sheet data
        for row in dataframe_to_rows(data, index=False, header=True):
            sheet.append(row)

        if metadata_name == 'TableInfos':
            # transpose TableInfos for easier reading
            transpose(sheet, min_row=1, max_row=1, min_col=1, max_col=21)
            transpose_row_to_col(sheet, min_row=1, max_row=1, min_col=1, max_col=21,target_cell_address=(3,1))
            transpose(sheet, min_row=2, max_row=2, min_col=1, max_col=21)
            transpose_row_to_col(sheet, min_row=2, max_row=2, min_col=1, max_col=21,target_cell_address=(3,2))
            sheet.delete_rows(1,2)
            sheet.column_dimensions['A'].width = 16
            sheet.column_dimensions['B'].width = 100
            start, stop = 1, 21
            for index, row in enumerate(sheet.iter_rows()):
                if start < index < stop:
                    for cell in row:
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        return sheet 
        
# **************************
# master control data
def masterControlData(data):
    # data object is TableInfos
    global controlInformationTables # all Tables evaluated
    global controlInformationTable  # current Table evaluated

    if not no_master:
    
        for ci in controlInformationTable:
            try:
                controlInformationTable[ci] = data[ci]
            except KeyError:
                controlInformationTable['lastRefreshDate'] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                break
                
    return
    
# ********************************************
# save master control file xlsx
# uses global var 
# controlInformationTable
# folder_name
# control_file
def masterControlFile(*fromProp):
    global controlInformationTable
    global controlInformationTables
    
    if not fromProp:
        
        output_file = folder_name + control_file
            
        if Path(output_file).is_file():
            try:
                controlbook = load_workbook(output_file)
                controlbook.active = controlbook['controlBook CBS']
                sheet = controlbook.active
                addHeader = False
                addValues = True

            except:
                controlbook = Workbook()    
                sheet = controlbook.create_sheet('controlBook CBS') # max first 31 chars
                addHeader = True
                addValues = True
                
        else:
            controlbook = Workbook()    
            sheet = controlbook.create_sheet('controlBook CBS') # max first 31 chars
            addHeader = True
            addValues = True
            
        lastrow = sheet.max_row+1
        
        if addHeader:
            # pd = pd.DataFrame(controlInformationTables)
            for col, val in enumerate(controlInformationTables.keys(), start=1):
                sheet.cell(row=1, column=col).value = val
    
        if addValues:
            for col, val in enumerate(controlInformationTables.values(), start=1):
                sheet.cell(lastrow, column=col).value = val
            
        try:
            controlbook.save(output_file) 
            
        except Exception as e:
            p(warning,'\nUnable to save master control workbook', output_file, "Do you have it open in Excel? The error message is", e)
            pass    

        controlInformationTable = {"Title":"","Updated":"","ShortTitle":"","Identifier":"","Summary":"","Modified":"","ReasonDelivery":"","Frequency":"","Period":"","RecordCount":"","lastRefreshDate":"","lastRefreshDateJson":"","lastRefreshDateCsv":"","lastRefreshDateExcel":""}
        controlInformationTables = {}


    return

# ****************************
# get endpoint using cbsodata
# on error with excel switch
def get_table_endpoint(p_identifier, endpoint, file_path, *workbook):
    
    global download_excel
            
    file_exists = False
    
    # does file exist and do we update and/or force?
    # otherwise skip
    if not force_download and not data_refresh:
        if download_csv:
            output_file = os.path.join(file_path, p_identifier+'-'+endpoint+ '.csv')
            if Path(output_file).is_file():
                file_exists = True
            
        if download_json:
           output_file = os.path.join(file_path, p_identifier+'-'+endpoint+ '.json')
           if Path(output_file).is_file():
                file_exists = True
                
        if download_excel:
            output_file = os.path.join(file_path, p_identifier+'-objects.xlsx')
            if Path(output_file).is_file():
                file_exists = True
                
    else:
        file_exists = False
        
    # do csv & excel & json
    if (not file_exists and download_data) and (download_csv or download_excel or download_json):

        try:                
            # put cbsodata respons in dataframe
            if endpoint == 'TableInfos':
                data = pd.DataFrame(cbsodata.get_meta(p_identifier, endpoint)) # pd.DataFrame(data_in, index=[0])
            else:
                data = pd.DataFrame(cbsodata.get_meta(p_identifier, endpoint))
            
        except Exception as e:
            p(warning, '\t\t\t\tUnable to retrieve object', endpoint, 'for table', p_identifier, '. The error message was\t\t\t\t', e)
            return

        if download_excel:
            workbook = convertTuple(workbook)            
            try:
                # just remove this sheet, leave the rest alone
                workbook.remove(workbook[endpoint[0:30]])
            except Exception as e:
                pass
                
            # create sheet for data
            sheet = workbook.create_sheet(endpoint[0:30]) # max first 31 chars
            workbook.active = workbook[endpoint[0:30]]
            sheet = workbook.active               

        if endpoint == 'DataProperties':
            # get extra Dimension endpoints
            # will be added to excel file as sheets
            data_np = data[['odata.type', 'Key']].to_numpy()
            for dimension in data_np:
                if dimension[0] == 'Cbs.OData.Dimension':
                    p(info, '\t\t\t\t\t... extra dimension ', dimension[1])
                    # do it again
                    get_table_endpoint(p_identifier, dimension[1], file_path, workbook if workbook else None)                   

        if download_csv:
            csv_data = save_data(data, file_path, p_identifier, endpoint,'csv')
            return csv_data
            
        if download_json:
            json_data = save_data(data.to_json(), file_path, p_identifier, endpoint, 'json')
            return json_data
            
        if download_excel:
            sheet = save_data(data, file_path, p_identifier, endpoint, sheet)
            return workbook
    else:
        p(warning, '\t\t\t\t\t... file exists but not updating')
    
    # only save excel output once
    return

# **********************************
# download data from table
# uses arg -f (default ./data/)
def get_table_meta(data, *endpoint):

    global download_excel
    
    p_identifier = data['Identifier']
        
    # table objects (endpoints for odata interface) excl. 'UntypedDataSet'
    objects_lst = ['DataProperties','Perioden','TableInfos','CategoryGroups','TypedDataSet']
    
    file_path=folder_name+p_identifier+'/'
    
    # when endpoints given only retrieve those
    retrieve_lst = []
    if len(endpoint)>0:
        for i in endpoint:
            retrieve_lst.append(i)
    else:
        retrieve_lst = objects_lst
        
    # export file based on argument + subfolders
    if not os.path.isdir(file_path):
        try:
            os.mkdir(file_path)
        except Exception as e:
            p(error,'Creating folder', file_path, 'failed with error', e, '. Do you have sufficient rights?')
            download_data = False

    if download_excel:
        p(info,'\t\texcel file name\t\t'+file_path+p_identifier+"-objects.xlsx")
        
        output_file = file_path+p_identifier+"-objects.xlsx"
        
        if Path(output_file).is_file():
            # file exists
            if data_refresh:        
                try:
                    # does the workbook exist? if so update sheets.
                    # don't mess with other sheets
                    workbook = load_workbook(file_path+p_identifier+"-objects.xlsx")
                except:                
                    # create new 
                    workbook = Workbook()
                    
            elif force_download: 
                # create new by -force
                workbook = Workbook()
                
            else:
                workbook = None    
                
        else:
            workbook = Workbook()    
    else:
        workbook = None
        
    if download_csv:
        p(info,'\t\tcsv file name(s)\t' + file_path+p_identifier+'-<object>.csv')
        
    if download_json:
        p(info,'\t\tjson file name(s)\t' + file_path+p_identifier+'-<object>.json')

    # save the data endpoints from list
    for i in retrieve_lst:
        p(info, '\t\t\t\t... performing update for', i)
        if i == 'TableInfos':
            return_value = get_table_endpoint(p_identifier, i, file_path, workbook)
        else:    
            return_value = get_table_endpoint(p_identifier, i, file_path, workbook)

    # saving excel
    if download_excel and not type(return_value) == None:        
        if not no_master:
            controlInformationTable['lastRefreshDateExcel'] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

        try:
            return_value.save(file_path+p_identifier+"-objects.xlsx")                    
        except AttributeError:
            if not data_refresh and not force_download:
                pass
            else:
                p(warning,'Updating excel file failed. If it exists you could use *-update* or *-force*.\n')
        except Exception as e:
            p(warning,'\nUnable to save workbook', file_path+p_identifier+"-objects.xlsx.", "Do you have it open in Excel?\n The error message is \n", e)
            pass
                    
    return return_value

# *************************************
# Convert (saved) text as json into list
def convertToList(data, datatype):
    
    try:
        if str(datatype) == 'CBSODATA':
            # create local copy of table list
            listObject = []
            try:
                with open(data, "rb") as f:
                    listObject = pickle.load(f) 
                f.close()                
            except Exception as e:
                listObject = []
                print(e)
            return listObject
            
        elif str(datatype) == 'CONTROL':
            # split for excel
            listObject = []
            try:
                listObject = data.split(',')                
            except Exception as e:
                listObject = []
                print(e)
            return listObject

        else:
            p(warning, 'Converting the file', file, 'of type', datatype, 'did not succeed. It seems like', datatype, 'is not implemented (yet).')
            listObject = []
            return listObject
            
    except Exception as e:
        p(error, 'While converting the file', file, 'to a list the following error occured:\n', e)
        listObject = []
        return listObject
        

# ****
# main get table list if no ID given
if not table_identifier:
    if len(glob.glob(folder_name + "cbs_all_tables.list")) > 0 and data_refresh:
        p(info, "A local copy (" + folder_name + "cbs_all_tables.json) has been found. Update is used, retrieving data from CBS odata endpoint and saving it as new local copy.")
        tables = cbsodata.get_table_list()
        # save as local copy
        pickle.dump( tables, open( folder_name + "cbs_all_tables.list", "wb" ) )
      
    elif len(glob.glob(folder_name + "cbs_all_tables.list")) > 0 and not data_refresh:     
        # we have a cbs_all_tables local copy
        p(info, "Using local copy of CBS table information in (" + folder_name + "cbs_all_tables.json).")
        tables = convertToList(folder_name + "cbs_all_tables.list", "CBSODATA")

    else:
        p(info, "There is no " + folder_name + "cbs_all_tables.list found. Retrieving data from CBS odata endpoint and saving it as new local copy.")
        tables = cbsodata.get_table_list()
        # save as local copy
        pickle.dump( tables, open( folder_name + "cbs_all_tables.list", "wb" ) ) 
        
else:
    # just this one table
    tables = []
    for i in table_identifier:
        tables.append(cbsodata.get_meta(i, 'TableInfos'))
    p(verbose, 'TableInfo downloaded from CBS for', len(tables), 'due to argument -i.')
    get_tables = len(tables)

# some vars for loop
all_tables = len(tables)    
end_record = all_tables
itable = 0   # count tables processed
itable_records = 0  # count nr of records in hits
number_of_hits = 0  # count search hits
result_list = []    # save table identifers
search_list = []  # save search arguments

if not start_record: 
    start_record = 0
else:
    p(verbose, 'Starting at record ', start_record)

p(verbose, 'Table list contains ' + str(all_tables) + ' tables ', 'starting at record '+ str(start_record) if start_record > 0 else '' )

# how much?
end_record = all_tables - start_record
    
if get_tables > 0: 
    p(info, 'Getting maximum of ' + str(get_tables) + ' tables due to argument -n or -i.')
    end_record = start_record + get_tables
    
if end_record > all_tables:
    end_record = all_tables
    p(verbose, 'End record is ' + str(end_record))
    
# Console messages 
p(info, 'Searching in ShortDescription for keyword(s) ' + str(search_arg) if search_arg else 'No search keywords (-s) given.')
p(verbose, 'Downloading data into folder ', folder_name)

if table_meta: p(verbose, 'Meta data will be downloaded...')

p(verbose, '\n--------------------')

if len(search_arg)>0:
    search_list = search_arg
else:
    search_list=[]

# start loop for all tables
for table in tables:
    itable+=1
    
    # loop until we get at the start record
    if itable >= start_record and itable <= end_record:
        datemodified = datetime.datetime.strptime(table['Modified'][0:10], '%Y-%m-%d')
        datemodified = datetime.datetime.date(datemodified)
        
        if modified_within == 'lastday':
            minmoddate = datetime.date.today() - datetime.timedelta(days=1)
        if modified_within == 'lastweek':
            minmoddate = datetime.date.today() - datetime.timedelta(days=7)
        if modified_within == 'lastmonth':
            minmoddate = datetime.date.today() - datetime.timedelta(days=30)
        if modified_within == 'lastyear':
            minmoddate = datetime.date.today() - datetime.timedelta(days=365)
            
        if modified_within:
            maxmoddate = datetime.date.today()
            if datemodified >= minmoddate and datemodified <= maxmoddate:
                modified_within_selection = True
                p(verbose, 'Table', table['Identifier'], 'modified on', datemodified, 'which is valid for the selection.')
            elif datemodified > maxmoddate:
                modified_within_selection = False
                p(verbose, 'Table has a modified date of', datemodified, 'which is in the future. Wauw.')
            elif datemodified < minmoddate: 
                modified_within_selection = False
                p(verbose, 'Table is last modified on', datemodified, 'which is too far in the past.')
        else:
            modified_within_selection = True
            
        if modified_within_selection:
            p(verbose, 'getting meta data', table['Identifier'])
            p(verbose, 'Identifier table ' + table['Identifier'] + '\nShortTitle table' + table['ShortTitle'])
            p(verbose, '\nShortDescription table' + table['ShortDescription'])            

            # search properties
            if search_list:
                for keyword in search_list:
                    isHit = False   # only one hit is needed
                    if not isHit:
                        if str(table['ShortDescription']).find(keyword) > 0:
                            number_of_hits += 1
                            isHit = True
                            p(info if not download_data else verbose, table['Identifier'], 'has in ShortDescription search item', keyword,'and is updated within the modified date selection criteria.' if modified_within_selection else keyword)
                            result_list.append(table)
                        else:
                            p(verbose, 'Search string not found', keyword)
                
            else: # no search parameters given
                p(verbose, 'Table', table['Identifier'], 'selected and added to the result list.')
                result_list.append(table)

    # stop searching
    if itable >= end_record:
        p(verbose, '--------------------')
        break

# any results? or just this one table
if len(result_list)>0:
    p(info, '\nNumber of tables to retrieve:', len(result_list))
    
    if len(result_list) > 60 and (not data_refresh or not force_download) and download_data and not table_prop:
        p(warning, "\nThis is (probably) a lot of data, please use -force -update to download. \nOr use parameter -m to download just the table information.")
    else:
        if download_data:

            for result in result_list:
                p(info, '\n\tCommencing retrievel of table data for', result['Identifier'], result_list.index(result)+1,'/',len(result_list))
                
                if table_prop:
                    get_table_meta(result, 'DataProperties')
                    # get info from tables
           
                if table_meta:
                    get_table_meta(result, 'TableInfos')
                        
                elif not table_meta and not table_prop:
                    # get all data and properties of table
                    get_table_meta(result)

                if not no_master:
                    masterControlData(result)
                    itable_records = itable_records + int(controlInformationTable['RecordCount'])
                    # send to master
                    controlInformationTable['lastRefreshDate'] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
                    controlInformationTables.update(controlInformationTable)    
                    masterControlFile()        


        else:
            p(warning, "Results not downloaded. Use argument '-d' for download or one of the file extensions (-csv, -xls, -json). When there are more then 60 tables also use -force -update.")
                    
    p(verbose, 'Finished retrieving results.')
    
else:
    p(warning,'No search strings and/or results found. Maybe something wrong with the filter?')

# stats for geeks
end = datetime.datetime.now() 
elapsed_time = (end - begin)    
p(silent, '\nTotal time passed ' + str(elapsed_time), ' seconds, which is ', elapsed_time/len(result_list) if len(result_list)>0 else 0, ' per table. In total', itable_records, '(unique) records in the selection.' )
