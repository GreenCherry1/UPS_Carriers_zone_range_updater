# UPS_Carriers_zone_range_updater
First task in DEXIMA internship
## Files
### Scripts
UPSScraper file downloads all .xls files of zones informations that are written in Carriers zone range file at sheet UPS wuth default names in xls_files folder
ConverterXlsToXlsx saves all .xls as .xlsx files  in xlsx_files folder
DataUpdater changes values of cells in Carriers zone range to zipcodes in the downloaded files
### Folders
xls_files saves all downloaded .xls files
xlsx_files saves all converted .xlsx files
### Text
requirements contains list of all used libraries
### xlsx
Carriers zone range contains data of regions
