# autocbs

Command line interface to the cbsodata modules with added support for creating excel (xlsx) and csv files from cbs data.

# usage

autocbs.py :: Download CBS data tables

Use argument -h (or omit any arguments) for the command line help
  
# examples

**1) autocbs.py -s jeugdzorg -f ./data/ -m -csv -force -update -mw lastyear**

The following arguments will get the tables list from the cbsodata endpoint and save it in the data folder with the name "cbs_all_tables.list". 
The data folder (-f) will be the current folder + './data/'. 
The ShortDescription will be searched (-s) for the keyword 'jeugdzorg' for tables which are updated (-mw) within the last year from today. 
The metadata (-m) of the table will be downloaded as a (-csv) file and placed into the folder ./data/<tableidentifier>/<tableidentifier>-TableInfos.csv.
The master control file is an excel file that will be placed in the './data/' folder with the name "get_data_control.xlsx" and contains the following entries from the selected tables meta data "Title, Updated, ShortTitle, Identifier, Summary, Modified, ReasonDelivery, Frequency, Period, RecordCount, lastRefreshDate, lastRefreshDateJson, lastRefreshDateCsv, lastRefreshDateExcel".
The arguments -force and -update will force an update of all files already downloaded in the './data/<tableidentifier>' folders.

A first run of the script with these arguments on a fresh install got 16 tables in the result list and took 08.54 seconds to perform.

**2) autocbs.py -s jeugdzorg -f ./data/ -m -csv -p -mw lastyear**

The second run queried for the same keyword with the same periode. This time the DataProperties (-p) of the tables are downloaded in csv format. The resulting downloads are added to the data folder. Depending on the properties of the DataProperties object and the values of the key "Cbs.OData.Dimension" these extra dimensions, which are available as an odata endpoint, are downloaded as well. The result is 56 csv files for a total of 418.541 bytes.

**3) autocbs.py -s jeugdzorg -f ./data/ -xls -mw lastyear -v info -force -update**

The third run did a download of the data for the same tables and created excel files, xlsx format, for them. A couple of files had more then 1.048.574 records, which is more then excel will handle. Total run time 44 minutes.

**4) autocbs.py -s jeugdzorg -f ./data/ -csv -mw lastyear -v info -force -update**

The fourth run did the data downloads for the same tables in csv format. The TypedDataset csv file (84135NED-TypedDataSet.csv) was 143.814KB big with 2.617.216 records. Download time 15:07. Totale records downloaded for the 16 tables was 6.336.214.

# remarks

-update:  When updating excel files only the sheets containing names corresponding to the table objects, including dimensions, are refreshed. Old data is removed and a new sheet with the same name is created. Other sheets will be intact.

-force:   When using argument force with excel files the old files will be overwritten.

-csv:     These files will be overwritten with either -update or -force.

-json:    These files will be overwritten with either -update or -force.
