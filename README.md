# PTV Visum TAG Databook Tool

## Summary

Example model and Python script to add tables from the TAG databook into PTV Visum as User-Defined Tables. The example model uses the January 2023 version of the databook (v1.20.2) with default Price and Value years (2010). 
Further tables are created to generate the parameters required for highway assignment and demand models.

## Example File

An example file is provided in this repository both before and after the script has been run. Users should note that this uses the following settings:
- Price Year = 2010
- Value Year = 2010
- Initial Forecast Year = 2010

Should alternative values be required, the parameters should be set in the databook spreadsheet and saved before the tool is run to ensure that the correct values are imported to Visum.

## Databook Importer

### Introduction

A Python script is provided in this repository that can be run to create and fill the data tables, or replace them if they already exist. This ensures that consistent values are used in all model runs, and gives transparency on the source of all data. There is built-in quality assurance functionality such as attributes that record:
- The date, time and user that imported the data
- The databook version
- The filepath of the databook imported
- The price, value and initial forecast years selected

### Prerequisites
There are no prerequisites for the tool to run, however if there are existing user-defined attributes or tables that have the same ID's, these will be removed and replaced.

There are a number of conditions and dependencies that should be filled by the user in order for the formulae to calculate properly, however.

#### Activity Pairs
An Activity Pairs attribute with ID AUC is required in order to look up the correct fuel cost (VOC) parameters for the demand model. If one does not exist, a new one is created by the script automatically that the user should fill. The expected values for this attribute are:
- CB
- CC
- CO
- LGV
- HGV

This allows the Value of Time (VOT) and occupancy (OCC) attributes to be calculated correctly.

#### Modes
Different modes should be created in the model for each AUC that is used, with the CODE attribute being set to CC, CB, CO, LGV, HGV. This is to ensure that the VOC fuel and non-fuel cost parameters are looked up correctly from the databook tables. This does not apply to public transport or active modes. 

Note: As per best practice, the non-fuel cost component of the VOC is set to 0 manually for business trips as it is assumed that this cost is not perceived by the transport user.


### databook.py
This is the script that performs the data import into Visum. There are a few options for how this can be run depending on the preference of the user.
Running the script for the first time will open a file selection dialog where the user should select the databook file that they want to import data from. The script will them import the data into the user-defined tables for use in demand and assignment models. A network attribute (DB_PATH) is created with the location of the databook. If the tool is re-run, this is the location the data will be read from again. 

#### Run Script Procedure
The tool can be run from the Visum procedure sequence with a Run Script procedure that either links to the databook.py script location, or the code can be manually copied into the window to avoid external files. 

For either method, the `src` folder that comes with the tool needs copying to either the location of the script or the Visum file as appropriate. 

#### Drag & Drop into Visum
The tool can alternatively be run by dragging and dropping the `databook.py` script into the Visum window. As above, the `src` folder needs to be in the same location as the script file.

#### Scripts > Run Script
The third option for running the tool is to navigate to Scripts > Run script file... and select the .py file. It is required that the `src` folder is located in the same folder as this Python script, but the version file and databook file do not necessarily need to be in the same location.

### Settings & Options
The two primary user inputs to the tool once it has run are the model year and time period.

#### Model Year
The model year is stored in the MODEL_YEAR network attribute and controls which values for forecast parameters are used. No further input is required to update parameters for future years. 

#### Model Time Period
The MODEL_TP parameter controls which values are used for the impedance calculation, as occupancies vary by time of day. Accepted values are:
- Average Weekday
- Average Weekend
- 7am-10am
- 10am-4pm
- 4pm-7pm
- 7pm-7am

It is recommended that users utilise Edit Attribute procedures in the procedure sequence to update this attribute as required for each assignment.

#### Other Options
Other attributes that can be changed are:
- OGV1_Proportion - the proportion of HGVs that are OGV1 (default 0.4)
- OGV2_Proportion - the proportion of HGVs that are OGV2 (default 0.6)
- HGV_VOT_Factor - the factor applied to the value of time for HGVs during routing (default 2.5)
- OVERRIDE_AVG_NET_SPEED - this controls whether the average network speed that is required for the assignment VOC parameters should be calculated from the network directly or whether a standard value should be used (in Perceived_VOC_int table)
- NO_OF_ITER_FOR_CONV - the iteration number of the assignment to take to calculate the average network speed if OVERRIDE_AVG_NET_SPEED is true


## Tables Created

- A1.1.1: Green Book Discount Rates
- A1.3.10:: Forecast fuel efficiency improvements
- A1.3.11:: Forecast fuel consumption parameters
- A1.3.12:: Forecast fuel cost parameters - Work
- A1.3.13: Fuel cost parameters - Non-Work
- A1.3.14: Non-fuel resource vehicle operating costs
- A1.3.15: Forecast non-fuel resource vehicle operating costs
- A1.3.16: Proportion of bus trips by car ownership, trip purpose and concessionary travel pass status
- A1.3.17: Proportion of bus trips by that would “not go” if bus not available
- A1.3.18: Value of the social impact per return bus trip
- A1.3.1a: Values of Working (Employers' Business) Time by Mode (£ per hour)
- A1.3.1b: Values of Non-Working Time by Trip Purpose (£ per hour)
- A1.3.1c: Parameter values for employers' business value of time by mode
- A1.3.1d: Values of Working (Employers' Business) Time by mode per person (distance banded)
- A1.3.2a: Forecast values of time per person - Working - Resource cost values (£ per hour)
- A1.3.2b: Forecast values of time per person - Non-Working - Resource cost values (£ per hour)
- A1.3.2c: Forecast values of time per person - Working - Perceived cost values (£ per hour)
- A1.3.2d: Forecast values of time per person - Non-Working - Perceived cost values (£ per hour)
- A1.3.2e: Forecast values of time per person - Working - Market price values (£ per hour)
- A1.3.2f: Forecast values of time per person - Non-Working - Market price values (£ per hour)
- A1.3.3a: Car occupancies per Vehicle Kilometre Travelled and per Trip by Journey Purpose
- A1.3.3b: Vehicle occupancies per Vehicle Kilometre Travelled
- A1.3.3c: Annual Percentage Change in Car Passenger Occupancy (% pa) up to 2036
- A1.3.4: Proportion of travel in work and non-work time
- A1.3.5: Market  Price Values of Time per Vehicle based on distance travelled
- A1.3.6: Market Price Values of Time per Vehicle based on distance travelled (£ per hour)
- A1.3.7: Fuel and Electricity Prices and Components
- A1.3.8: Fuel consumption parameter values
- A1.3.9: Proportion of cars, LGV & other vehicle kilometres using petrol, diesel or electricity
- Perceived_VOC_final: Final Vehicle Operating Costs
- Perceived_VOC_int: Interim Vehicle Operating Costs
- Perceived_VOT_final: Final Perceived Value of Time - Goods Vehicle Aggregated
- Perceived_VOT_int: Interim Perceived Value of Time - Goods Vehicle Disaggregated
- UDAs_for_Impedance: Values used in the assignment for impedance

## Attributes Created

### Activity Pairs
- AUC - If this doesn't exist it is created to match assignment user classes with activity pairs

### Network
- DB_IMPORT_DATETIME - The date and time the tool was last run
- DB_USER - The username of the user who last ran the tool
- DB_PATH - The filepath of the databook
- DB_VERSION - The version number of the databook
- DB_PRICE_YEAR - The price year set in the databook
- DB_INITIAL_FORECAST_YEAR - the initial forecast year set in the databook
- DB_VALUE_YEAR - the value year set in the databook
- INDIRECT_TAX_CORRECTION - the value of the indirect tax correction divisor
- MODEL_YEAR - the year of the model 
- MODEL_TP - the time period of the model
- NO_OF_ITER_FOR_CONV - the number of the assignment iteration that should be taken for the average speed calculation
- OGV1_PROPORTION - the proportion of HGVs that are OGV1
- OGV2_PROPORTION - the proportion of HGVs that are OGV2
- HGV_VOT_FACTOR - the value of time perception factor for HGV routing
- OVERRIDE_AVG_NET_SPEED - a flag for whether hardcoded average network speeds should be taken or whether it should be calculated from the assignment
- CB_IMP_DIST - the distance coefficient for CB in the impedance
- CC_IMP_DIST - the distance coefficient for CC in the impedance
- CO_IMP_DIST - the distance coefficient for CO in the impedance
- LGV_IMP_DIST - the distance coefficient for LGV in the impedance
- HGV_IMP_DIST - the distance coefficient for HGV in the impedance
- CB_IMP_TIME - the time coefficient for CB in the impedance
- CC_IMP_TIME - the time coefficient for CC in the impedance
- CO_IMP_TIME - the time coefficient for CO in the impedance
- LGV_IMP_TIME - the time coefficient for LGV in the impedance
- HGV_IMP_TIME - the time coefficient for HGV in the impedance
- CB_IMP_TOLL - the toll coefficient for CB in the impedance
- CC_IMP_TOLL - the toll coefficient for CC in the impedance
- CO_IMP_TOLL - the toll coefficient for CO in the impedance
- LGV_IMP_TOLL - the toll coefficient for LGV in the impedance
- HGV_IMP_TOLL - the toll coefficient for HGV in the impedance


## Advantages

- All data available and easily checked
- Future year updates are much easier
- For highway assignment, assumed network speed can be directly calculated from the assignment and used in the GenCost equations exactly rather than requiring calibration
- Ensures consistency between values in demand and assignment models

## QA and Checks

- Cross-checked final assignment values against known and validated models 
- Manually reviewed each table
- Used by PTV in internal projects 

## Disclaimer
Note that PTV cannot accept liability for any loss or damages arising from the use of this tool or the values generated by it. Users should complete their own tests and checks to ensure the values produced are those that they expect.
