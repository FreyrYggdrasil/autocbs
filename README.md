# autocbs

Command line interface to the cbsodata modules with added support for creating excel (xlsx) and csv files from cbs data.

# usage

    autocbs.py :: Download CBS data tables

    arguments:
    -h                      this help
    -d                      download data for table (in -f). Implies -csv, -json
                            and -xls if none of them are given
    -f <folder>             folder name for data download, Identifier is added to
                            the path, ./data/is the default
    -s <string,>            search for keywords in table ShortDescription (can be
                            comma seperated)
    -id <identifier,>       table(s) to download using on TableInfos.Identifier
    -v <level>              stdout output level (less->more) silent, critical,
                            error, warning, info, verbose, allmsg
    -n <nr>                 maximum tables to get (use this while testing)
    -b <nr>                 start at record (use this while testing)
    -m                      get meta data (TableInfos) of the selected table(s)
    -nm                     do NOT maintain master excel (get_data_control.xlsx)
                            with table info
    -p                      get the DataProperties of the table
    -csv                    save files as csv
    -force                  force download of large result set (will still skip
                            excel sheet TypedDataset when records > 1.000.000)
    -update                 update already downloaded tables
    -xls                    download/update excel file with table objects (will skip
                            TypedDataSet for records > 1.000.000)
    -json                   update json files
    -mw                     table modifed within lastday, lastweek, lastmonth
                            or lastyear

# Data structure
Following list is composed from reverse engineering the data downloads :punch:

## TableListInfos

The information present in a record retrieved with cbsodata.get_table_list (http://opendata.cbs.nl/ODataCatalog/Tables).

Contains the attributes
1..7  | 8..14 | 15..21 | 22..26
------------- | ------------- | ------------- | -------------
"Updated" | "Modified" | "Catalog" | "DefaultSelection"
"ID" | "MetaDataModified" | "Frequency" | "GraphTypes"
"Identifier" | "ReasonDelivery" | "Period" | "RecordCount"
"Title" | "ExplanatoryText" | "SummaryAndLinks" | "ColumnCount"
"ShortTitle" | "OutputStatus" | "ApiUrl" | "SearchPriority"
"ShortDescription" | "Source" | "FeedUrl" |
"Summary" | "Language" | "DefaultPresentation" |

## TableInfos

Contains the attributes 
1..6  | 7..12 | 13..18 | 19..21
------------- | ------------- | ------------- | -------------
"ID" | "ReasonDelivery" | "ShortDescription" | "Source"
"Title" | "ExplanatoryText" | "Description" | "MetaDataModified"
"ShortTitle" | "Language" | "DefaultPresentation" | "SearchPriority"
"Identifier" | "Catalog" | "DefaultSelection" |
"Summary" | "Frequency" | "GraphTypes" |
"Modified" | "Period" | "OutputStatus" |

## DataProperties

Contains the attributes 
1..4  | 5..8 | 9..12 | 13..15
------------- | ------------- | ------------- | -------------
"odata.type" | "Type" | "MapYear" | "Decimals"
"ID" | "Key" | "ReleasePolicy" | "Default"
"Position" | "Title" | "Datatype" | "PresentationType"
"ParentID" | "Description" | "Unit" |

The _DataProperties."odata.type"="Cbs.OData.Dimension"_ may contain a seperate _"Dimension"_ type () which points to specific dimensions valid for the Table, e.g. "JeugdhulpInNatura", "Jeugdbescherming" and "Jeugdreclassering" for table 84137NED (-> get it with autocbs.py -d -p -xls -i 84137NED). These are downloaded as well.

The key _DataProperties."odata.type"="Cbs.OData.Topic"_ can have type _"Topic"_ with as _"Key"_ a column for the Table, e.g. "JongerenMetJeugdhulpInNatura_1", "JeugdbeschEnJeugdhMetVerblInNat_4" and "JongerenMetJeugdreclassering_5" for table 82972NED (-> get it with autocbs.py -d -p -xls -i 82972NED).  

The key _DataProperties."odata.type"="Cbs.OData.TimeDimension"_ can have type _"TimeDimension"_ with as _"Key"_ value "Perioden" which will point to an odata endpoint for the specific periods covered by the Table.  

## Perioden

Contains the attributes
1..4  |
------------- |
"Key" |
"Title" |
"Description" |
"Status" |

## CategoryGroups

Information about the grouping of the extra Dimensions.

Contains the attributes
1..5  |
------------- |
"ID"
"DimensionKey"
"Title"
"Description"
"ParentID"

## TypedDataset

The data itself. Contains all records for the table. Layout may differ on a table basis.

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

- -update:  When updating excel files only the sheets containing names corresponding to the table objects, including dimensions, are refreshed. Old data is removed and a new sheet with the same name is created. Other sheets will be intact.
- -force:   When using argument force with excel files the old files will be overwritten.
- -csv:     These files will be overwritten with either -update or -force.
- -json:    These files will be overwritten with either -update or -force.
