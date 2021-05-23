# PTV Visum TAG Databook Tables

## Summary

Example model and Python script to add tables from the TAG databook into PTV Visum as User-Defined Tables. The example model uses the July 2020 version of the databook with default Price and Value years (2010). 

## Example File

An example file is provided in this repository that contains each table from the databook. Users should note that this uses the following settings:
- Price Year = 2010
- Value Year = 2010
- Initial Forecast Year = 2010

Should alternative values be requied, the parameters should be set in the databook spreadsheet and saved before the tool is run to ensure that the correct values are imported to Visum.

## Databook Importer

### Introduction

A Python script is provided in this repository that can be run from inside Visum to create and fill the data tables, or update them if they already exist. This ensures that consistent values are used in all model runs, and gives transparency on the source of all data. There is built-in quality assurance functionality such as attributes that record:
- The date, time and user that imported the data
- The databook version
- The filepath of the databook imported
- The price and value years selected

### Import_Databook.py
This is the script that performs the data import into Visum. It can be run by either dragging and dropping the .py file into the Visum window, or by navigating to Scripts > Run script file... and selecting the .py file.

This will open a file selection dialog where the user should select the databook file that they want to import data from. The script will them import the data into the user-defined tables for use in demand and assignment models.


### Advantages

- All data available and easily checked
- Future year updates are much easier
- For highway assignment, assumed network speed can be directly calculated from the assignment and used in the GenCost equations exactly rather than requiring calibration
- Ensures consistency between values in demand and assignment models

### Future Extensions
- Processing the raw tables into usable parameters for demand and assignment models
