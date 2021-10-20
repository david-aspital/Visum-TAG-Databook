import os
import sys
sys.path.append(f"{os.path.dirname(os.path.realpath(__file__))}\\src")
import openpyxl
import wx
import datetime
import pandas as pd
import numpy as np
import traceback


def create_attributes(db_path):

    # Price, initial forecast and value years
    years = pd.read_excel(db_path, sheet_name='User Parameters', skiprows=9, usecols='L', engine='openpyxl').dropna(axis=0).reset_index(drop=True)
    price_year = years.Value[0]
    initial_year = years.Value[1]
    value_year = years.Value[2]
    print('Test')

    # Attributes to be added or updated and their types
    atts = {'DB_IMPORT_DATETIME' : (5, datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S')),
            'DB_USER' : (5, os.getlogin()),
            'DB_PATH' : (5, db_path),
            'DB_VERSION' : (5, pd.read_excel(db_path, sheet_name='Cover', skiprows=2, engine='openpyxl')['TAG Data Book'][0]),
            'DB_PRICE_YEAR' : (1, price_year), 
            'DB_INITIAL_FORECAST_YEAR' : (1, initial_year), 
            'DB_VALUE_YEAR' : (1, value_year),
            'INDIRECT_TAX_CORRECTION' : (2, pd.read_excel(db_path, sheet_name='A1.3.1', skiprows=14, engine='openpyxl')['Unnamed: 2'][0])}

    # Try to add attribute (ignore if already exists), then update value
    for att, value in atts.items():
        try:
            if value[0] == 2:
                Visum.Net.AddUserDefinedAttribute(att, att, att, value[0], 4)
            else:
                Visum.Net.AddUserDefinedAttribute(att, att, att, value[0])
        except:
            pass
        Visum.Net.SetAttValue(att, value[1])


def file_select_dlg(message, wildcard):
    with wx.FileDialog(parent=None, message=message, wildcard=wildcard,
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:

        if dlg.ShowModal() == wx.ID_CANCEL:
            exit(0)
        pathname = dlg.GetPath()
    return pathname


def create_fill_udt(df, name, comment):
    # Function to create or update user-defined table in Visum
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
        exists = True
    except:
        exists = False
    if not exists:
        dtypes = df.dtypes.to_dict()   
        udt = Visum.Net.AddTableDefinition(name)
        udt.SetAttValue('Comment', comment)
        udt.AddMultiTableEntries(list(range(1,len(df)+1)))
    for col in df.columns:
        uda_id = col.replace(" ", "_")
        if not exists:
            typ = dtypes[col]
            if typ == 'int' or typ == 'int64':
                typ = 1
            elif typ == 'float64':
                typ = 2
            elif typ == 'O':
                typ = 5
            else:
                raise ValueError(f'Unsupported type: {typ}')
            if typ == 'float64':
                udt.TableEntries.AddUserDefinedAttribute(uda_id, col, col, typ, 4, canBeEmpty=1)
            else:
                udt.TableEntries.AddUserDefinedAttribute(uda_id, col, col, typ)
        udt.TableEntries.SetMultiAttValues(uda_id, tuple(zip(range(1, len(df)+1), df[col].tolist())))


def a1_1_1(db_path):
    name = 'A1.1.1'
    comment = 'Green Book Discount Rates'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=23, engine='openpyxl', usecols='B,D').dropna()
    df['Years from current year'] = df['Unnamed: 1']
    df = df[['Years from current year', 'Discount rate']]
    df[['Lower Bound', 'Upper Bound']] = df['Years from current year'].str.split('-', expand=True)
    df['Upper Bound'].fillna(999999, inplace=True)
    df['Lower Bound'] = df['Lower Bound'].str.replace(' and over', '')
    df[['Lower Bound', 'Upper Bound']] = df[['Lower Bound', 'Upper Bound']].astype(int)
    df = df[['Years from current year', 'Lower Bound', 'Upper Bound', 'Discount rate']]
    create_fill_udt(df, name, comment)


def a_1_3_1(db_path):
    name = 'A1.3.1'
    comment = "Values of Working (Employers' Business) Time by Mode (£ per hour)"
    names = ['Mode', 'Factor Cost', 'Perceived Cost', 'Market Price']
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='A,D:F', nrows=14, header=None, names=names, engine='openpyxl')
    create_fill_udt(df, f'{name}a', comment)

    comment = 'Values of Non-Working Time by Trip Purpose (£ per hour)' 
    names = ['Trip Purpose', 'Factor Cost', 'Perceived Cost', 'Market Price']
    df = pd.read_excel(db_path, sheet_name=name, skiprows=44, usecols='A,D:F', nrows=2, header=None, names=names, engine='openpyxl')
    create_fill_udt(df, f'{name}b', comment)

    comment = "Parameter values for employers' business value of time by mode"
    df = pd.read_excel(db_path, sheet_name=name, skiprows=38, nrows=8, usecols='H:J', engine='openpyxl')
    create_fill_udt(df, f'{name}c', comment)

    comment = "Values of Working (Employers' Business) Time by mode per person (distance banded)" 
    names = ['Mode', 'Resource Cost', 'Perceived Cost', 'Market Price']
    df = pd.read_excel(db_path, sheet_name=name, skiprows=51, usecols='H:K', header=None, names=names, nrows=8, engine='openpyxl')
    df['Distance Band'] = df['Mode'].str.split(" ").str[-1]
    df['Mode'] = df.apply(lambda x: x['Mode'].replace(" "+x['Distance Band'], ""), axis=1)
    df[['Lower Bound', 'Upper Bound']] = df['Distance Band'].str.split('-', expand=True)
    df['Lower Bound'] = df['Lower Bound'].str.replace('km', '').str.replace('+', '')
    df['Upper Bound'] = df['Upper Bound'].str.replace('km', '')
    df['Upper Bound'].fillna(999999, inplace=True)
    df[['Lower Bound', 'Upper Bound']] = df[['Lower Bound', 'Upper Bound']].astype(int)
    df = df[['Mode', 'Distance Band', 'Lower Bound', 'Upper Bound', 'Resource Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}d', comment)


def a1_3_2(db_path):
    name = 'A1.3.2'
    comment = 'Forecast values of time per person - Working - Resource cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, D:Q', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, D:Q', header=None, names=header, engine='openpyxl')
    df['Year'] = df.Year.astype(int)
    create_fill_udt(df, f'{name}a', comment)

    comment = 'Forecast values of time per person - Non-Working - Resource cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, R:S', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, R:S', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}b', comment)

    comment = 'Forecast values of time per person - Working - Perceived cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, T:AG', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, T:AG', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}c', comment)

    comment = 'Forecast values of time per person - Non-Working - Perceived cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AH:AI', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, AH:AI', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}d', comment)

    comment = 'Forecast values of time per person - Working - Market price values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AJ:AW', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, AJ:AW', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}e', comment)

    comment = 'Forecast values of time per person - Non-Working - Market price values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AX:AY', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, nrows=80, usecols='B, AX:AY', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}f', comment)


def a1_3_3(db_path):
    name = 'A1.3.3'
    comment = 'Car occupancies per Vehicle Kilometre Travelled and per Trip by Journey Purpose'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='A,D:J', nrows=1, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header = [x.strip() for x in header]
    header[0] = 'Journey Purpose'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='A,D:J', nrows=4, header=None, names=header, engine='openpyxl')
    df = df.melt(id_vars='Journey Purpose', value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Occupancy Per VehKm')
    df2 = pd.read_excel(db_path, sheet_name=name, skiprows=31, usecols='A,D:J', nrows=4, header=None, names=header, engine='openpyxl')
    df2 = df2.melt(id_vars='Journey Purpose', value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Occupancy Per Trip')
    df3 = df.merge(df2)
    create_fill_udt(df3, f'{name}a', comment)

    comment = 'Vehicle occupancies per Vehicle Kilometre Travelled'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=36, usecols='A,B,H:J', nrows=1, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header = [x.strip() for x in header]
    header[0] = 'Vehicle Type'
    header[1] = 'Journey Purpose'
    header[2] = 'Average '+header[2]
    for i in range(3,5):
        header[i] = header[i]+' Average'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=38, usecols='A,B,H:J', nrows=7, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df = df.melt(id_vars=['Vehicle Type', 'Journey Purpose'], value_vars=['Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Occupancy Per VehKm')
    create_fill_udt(df, f'{name}b', comment)

    comment = 'Annual Percentage Change in Car Passenger Occupancy (% pa) up to 2036'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=48, usecols='A,D:J', nrows=1, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header = [x.strip() for x in header]
    header[0] = 'Journey Purpose'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=49, usecols='A,D:J', nrows=2, header=None, names=header, engine='openpyxl')
    df = df.melt(id_vars='Journey Purpose', value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average', 'Weekend', 'All Week'], var_name='Time Period', value_name='Change in Car Passenger Occupancy')
    create_fill_udt(df, f'{name}c', comment)

def a1_3_4(db_path):
    name = 'A1.3.4'
    comment = 'Proportion of travel in work and non-work time'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='A,B,D:J', nrows=1, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header = [x.strip() for x in header]
    header[0] = 'Mode'
    header[1] = 'Journey Purpose'
    header[6] = 'Average Weekday'
    header[7] = 'Weekend Average'
    header[8] = 'All Week Average'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=25, usecols='A,B,D:J', nrows=7, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df = df.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Distance Travelled by Vehicles')
    df2 = pd.read_excel(db_path, sheet_name=name, skiprows=33, usecols='A,B,D:J', nrows=12, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df2 = df2.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Distance Travelled by Occupants')
    df3 = pd.read_excel(db_path, sheet_name=name, skiprows=25, usecols='A,B,K:Q', nrows=7, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df3 = df3.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Vehicle Trips')
    df4 = pd.read_excel(db_path, sheet_name=name, skiprows=33, usecols='A,B,K:Q', nrows=12, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df4 = df4.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Person Trips')
    
    df5 = df.merge(df2, how='outer').merge(df3, how='outer').merge(df4, how='outer').fillna('')
    create_fill_udt(df5, f'{name}', comment)

def a1_3_5(db_path):
    name = 'A1.3.5'
    comment = 'Market  Price Values of Time per Vehicle based on distance travelled'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=25, usecols='A,B,D:J', nrows=1, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header = [x.strip() for x in header]
    header[0] = 'Mode'
    header[1] = 'Journey Purpose'
    header[6] = 'Average Weekday'
    header[7] = 'Weekend Average'
    header[8] = 'All Week Average'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=25, usecols='A,B,D:J', nrows=12, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df = df.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Market Price Value of Time Per Vehicle')
    df['Mode'] = np.where(df.Mode == 'PSV ', 'PSV (Occupants)', np.where(df.Mode == '(Occupants)', 'PSV (Occupants)', df.Mode))
    create_fill_udt(df, f'{name}', comment)

def a1_3_6(db_path):
    name = 'A1.3.6'
    comment = 'Market Price Values of Time per Vehicle based on distance travelled (£ per hour)'
    df = pd.read_excel(db_path, sheet_name=name, nrows=80, skiprows=23, header=[0,1,2,3], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col else ','.join(col).strip() for col in df.columns.values]
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Market Price Value of Time Per Vehicle')
    df[['Day Type', 'Time Period', 'Mode', 'Journey Purpose']] = df['Variables'].str.split(',', expand=True)
    df = df[['Year', 'Day Type', 'Time Period', 'Mode', 'Journey Purpose', 'Market Price Value of Time Per Vehicle']]
    df['Time Period'] = np.where(df['Time Period'].str.contains('Unnamed'), 'All', df['Time Period'])
    create_fill_udt(df, f'{name}', comment)

def a1_3_7(db_path):
    name = 'A1.3.7'
    comment = 'Fuel and Electricity Prices and Components'
    df = pd.read_excel(db_path, sheet_name=name, nrows=91, skiprows=23, header=[0, 1,2,3], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col else ','.join(col).strip() for col in df.columns.values]
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Component', 'Fuel Type', 'Mode', 'Unit']] = df['Variables'].str.split(',', expand=True)
    df['Mode'] = np.where(df.Mode.str.contains('Unnamed'), 'All', df.Mode)
    df = df[['Year', 'Component', 'Fuel Type', 'Mode', 'Unit', 'Value']]
    create_fill_udt(df, f'{name}', comment)


def a1_3_8(db_path):
    name = 'A1.3.8'
    comment = 'Fuel consumption parameter values'
    df = pd.read_excel(db_path, sheet_name=name, nrows=7, skiprows=24, header=[0, 1], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['Vehicle Category' if 'Vehicle' in col[1] else ' '.join(col).strip() for col in df.columns.values]
    df.columns = ['Max speed kph' if 'Max speed' in col else col for col in df.columns.values]
    df2 = pd.read_excel(db_path, sheet_name=name, nrows=4, skiprows=35, header=None,  names=df.columns.values.tolist(), engine='openpyxl', usecols='A,D:I')
    df2.fillna(0, inplace=True)
    df3 = df.append(df2, ignore_index=True)
    df3.rename({'Parameters a':'Param_a', 'Parameters b':'Param_b', 'Parameters c':'Param_c', 'Parameters d':'Param_d'}, axis=1, inplace=True)
    create_fill_udt(df3, f'{name}', comment)

def a1_3_9(db_path):
    name = 'A1.3.9'
    comment = 'Proportion of cars, LGV & other vehicle kilometres using petrol, diesel or electricity'
    df = pd.read_excel(db_path, sheet_name=name, nrows=47, skiprows=23, header=[0, 1], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col[1] else ','.join(col).strip() for col in df.columns.values]
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Mode', 'Fuel Type',]] = df['Variables'].str.split(',', expand=True)
    df = df[['Year', 'Mode', 'Fuel Type', 'Value']]
    create_fill_udt(df, f'{name}', comment)

def a1_3_10(db_path):
    name = 'A1.3.10'
    comment = 'Forecast fuel efficiency improvements'
    df = pd.read_excel(db_path, sheet_name=name, nrows=44, skiprows=24, header=[0, 1, 2], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['ToYear' if 'Year' in col[2] else 'FromYear' if 'From' in col[2] else ';'.join(col).strip() for col in df.columns.values]
    df = df.loc[:,~df.columns.duplicated()]
    df = df.melt(id_vars=['FromYear', 'ToYear'], var_name='Variables', value_name='Value')
    df[['Change', 'Mode', 'Fuel Type',]] = df['Variables'].str.split(';', expand=True)
    df['Change'] = np.where(df.Change.str.contains('Cumulative'), 'Cumulative', 'Annual')
    df['FromYear'] = df.FromYear.str.replace(' to', '').astype(int)
    df = df[['FromYear', 'ToYear', 'Change', 'Mode', 'Fuel Type', 'Value']]
    create_fill_udt(df, f'{name}', comment)

def a1_3_11(db_path):
    name = 'A1.3.11'
    comment = 'Forecast fuel consumption parameters'
    df = pd.read_excel(db_path, sheet_name=name, nrows=80, skiprows=23, header=[0, 1], engine='openpyxl')
    df.drop([('Unnamed: 0_level_0', 'Unnamed: 0_level_1'), ('Vehicle Category', 'Year.1')], axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col[1] else ';'.join(col).strip() for col in df.columns.values]
    df.fillna(0, inplace=True)
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Vehicle Type', 'Parameter',]] = df['Variables'].str.split(';', expand=True)
    df['Parameter'] = 'Param_'+df.Parameter
    df['Vehicle Type'] = df['Vehicle Type'].str.replace('Car1', 'Car')
    df = df.pivot_table(values='Value', index=['Year', 'Vehicle Type'], columns='Parameter').reset_index().sort_values(['Vehicle Type', 'Year'])
    df = df[['Year', 'Vehicle Type', 'Param_a', 'Param_b', 'Param_c', 'Param_d']]
    create_fill_udt(df, f'{name}', comment)


def a1_3_12(db_path):
    name = 'A1.3.12'
    comment = 'Forecast fuel cost parameters - Work'
    df = pd.read_excel(db_path, sheet_name=name, nrows=80, skiprows=23, header=[0, 1, 2], engine='openpyxl')
    df.drop([('Unnamed: 0_level_0', 'Unnamed: 0_level_1', 'Unnamed: 0_level_2'), ('Unnamed: 2_level_0', 'Unnamed: 2_level_1', 'Year')], axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col[2] else ';'.join(col).strip() for col in df.columns.values]
    df.fillna(0, inplace=True)
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Vehicle Type', 'Fuel Type', 'Parameter']] = df['Variables'].str.split(';', expand=True)
    df['Parameter'] = 'Param_'+df.Parameter
    df['Fuel Type'] = df['Fuel Type'].str.replace('Car1', 'Car')
    df = df.pivot_table(values='Value', index=['Year', 'Vehicle Type', 'Fuel Type'], columns='Parameter').reset_index().sort_values(['Vehicle Type', 'Fuel Type', 'Year'])
    df.drop('Param_d.1', axis=1, inplace=True, errors='ignore')
    df = df[['Year', 'Vehicle Type', 'Fuel Type', 'Param_a', 'Param_b', 'Param_c', 'Param_d']]
    df['Vehicle Type'] = np.where(df['Vehicle Type']=='OGV', np.where(df['Fuel Type'].str.contains('OGV1'), 'OGV1', 'OGV2'), df['Vehicle Type'])
    #TODO replace below with dictionary normalisation
    df['Vehicle Type'] = np.where(df['Vehicle Type']=='Cars', 'Car', df['Vehicle Type'])
    df['Fuel Type'] = df.apply(lambda x: x['Fuel Type'].replace(x['Vehicle Type'], ""), axis=1)
    df['Fuel Type'] = df['Fuel Type'].str.strip()
    create_fill_udt(df, f'{name}', comment)

def a1_3_13(db_path):
    name = 'A1.3.13'
    comment = 'Fuel cost parameters - Non-Work'
    df = pd.read_excel(db_path, sheet_name=name, nrows=80, skiprows=23, header=[0, 1, 2], engine='openpyxl')
    df.drop([('Unnamed: 0_level_0', 'Unnamed: 0_level_1', 'Unnamed: 0_level_2'), ('Unnamed: 2_level_0', 'Unnamed: 2_level_1', 'Year')], axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col[2] else ';'.join(col).strip() for col in df.columns.values]
    df.fillna(0, inplace=True)
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Vehicle Type', 'Fuel Type', 'Parameter']] = df['Variables'].str.split(';', expand=True)
    df['Parameter'] = 'Param_'+df.Parameter
    df['Fuel Type'] = df['Fuel Type'].str.replace('Car1', 'Car')
    df = df.pivot_table(values='Value', index=['Year', 'Vehicle Type', 'Fuel Type'], columns='Parameter').reset_index().sort_values(['Vehicle Type', 'Fuel Type', 'Year'])
    df.drop('Param_d.1', axis=1, inplace=True, errors='ignore')
    df = df[['Year', 'Vehicle Type', 'Param_a', 'Param_b', 'Param_c', 'Param_d']]
    create_fill_udt(df, f'{name}', comment)

def a1_3_14(db_path):
    name = 'A1.3.14'
    comment = 'Non-fuel resource vehicle operating costs'
    names =  ['Vehicle Type', 'Fuel Type', 'Param_a1', 'Param_b1']
    df = pd.read_excel(db_path, sheet_name=name, nrows=13, skiprows=25, engine='openpyxl', index_col=[0,1]).reset_index()
    df.dropna(axis=1, inplace=True)
    df.columns = names
    df[['Trip Purpose', 'Fuel Type']] = df['Fuel Type'].str.split(' ', expand=True)
    df['Fuel Type'] = df.apply(lambda row: 'Non-electric' if ((pd.isna(row['Fuel Type'])) & (row['Vehicle Type']=='LGV')) else 'All', axis=1)
    df['Fuel Type'] = df['Fuel Type'].str.replace('Electic', 'Electric')
    df = df[['Vehicle Type', 'Trip Purpose', 'Fuel Type', 'Param_a1', 'Param_b1']]
    create_fill_udt(df, f'{name}', comment)

def a1_3_15(db_path):
    name = 'A1.3.15'
    comment = 'Forecast non-fuel resource vehicle operating costs'
    df = pd.read_excel(db_path, sheet_name=name, nrows=36, skiprows=25, header=None, usecols='B,D:I', engine='openpyxl')
    headers = ['Year', 'Work;Car;Param_a1', 'Work;Car;Param_b1','Non-work;Car;Param_a1','Non-work;Car;Param_b1','Average;Car;Param_a1', 'Average;Car;Param_b1']
    df.columns = headers
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Value')
    df[['Trip Purpose', 'Vehicle Type', 'Parameter']] = df['Variables'].str.split(';', expand=True)
    df = df.pivot_table(values='Value', index=['Year', 'Trip Purpose', 'Vehicle Type'], columns='Parameter').reset_index().sort_values(['Trip Purpose', 'Vehicle Type', 'Year'])
    create_fill_udt(df, f'{name}', comment)

def a1_3_16(db_path):
    name = 'A1.3.16'
    comment = 'Proportion of bus trips by car ownership, trip purpose and concessionary travel pass status'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=27, header=None, usecols='A:I', engine='openpyxl')
    headers = ['HH car ownership', 'Trip Purpose', 'Concessionary pass status', 'London Boroughs', 'Metropolitan built-up areas', 'Large and medium urban areas', 'Small urban and rural (<10k popn)', 'All areas (exc London)', 'All areas (inc London)']
    df.columns = headers
    df = df.melt(id_vars = ['HH car ownership', 'Trip Purpose', 'Concessionary pass status'], var_name = 'Area', value_name='Value')
    df.dropna(axis=0, inplace=True)
    create_fill_udt(df, f'{name}', comment)

def a1_3_17(db_path):
    name = 'A1.3.17'
    comment = 'Proportion of bus trips by that would “not go” if bus not available'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=27, header=None, usecols='A:D', engine='openpyxl')
    headers = ['HH car ownership', 'Trip Purpose', 'Concessionary pass status', 'Proportion not go']
    df.columns = headers
    df.dropna(axis=0, inplace=True)
    create_fill_udt(df, f'{name}', comment)

def a1_3_18(db_path):
    name = 'A1.3.18'
    comment = 'Value of the social impact per return bus trip'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=27, header=None, nrows = 2, usecols='A,D', engine='openpyxl')
    headers = ['Concessionary travel pass status', 'Value']
    df.columns = headers
    create_fill_udt(df, f'{name}', comment)

if __name__ == '__main__':
    app = wx.App()
    wildcard = "Excel Files(*.xlsm; *.xlsx)|*.xlsm;*.xlsx|" "All files (*.*)|*.*"
    db_path = file_select_dlg("Please select TAG Databook file...", wildcard)
    num_tables = 19

    try:
        progress_dlg = wx.ProgressDialog("Importing Tables", "Importing tables from databook...", num_tables+1, style=wx.PD_APP_MODAL | wx.PD_SMOOTH | wx.PD_AUTO_HIDE)
        create_attributes(db_path)
        progress_dlg.Update(1, "Importing Table A1.1.1...")
        a1_1_1(db_path)
        progress_dlg.Update(2, "Importing Table A1.3.1...")
        a_1_3_1(db_path)
        progress_dlg.Update(3, "Importing Table A1.3.2...")
        a1_3_2(db_path)
        progress_dlg.Update(4, "Importing Table A1.3.3...")
        a1_3_3(db_path)
        progress_dlg.Update(5, "Importing Table A1.3.4...")
        a1_3_4(db_path)
        progress_dlg.Update(6, "Importing Table A1.3.5...")
        a1_3_5(db_path)
        progress_dlg.Update(7, "Importing Table A1.3.6...")
        a1_3_6(db_path)
        progress_dlg.Update(8, "Importing Table A1.3.7...")
        a1_3_7(db_path)
        progress_dlg.Update(9, "Importing Table A1.3.8...")
        a1_3_8(db_path)
        progress_dlg.Update(10, "Importing Table A1.3.9...")
        a1_3_9(db_path)
        progress_dlg.Update(11, "Importing Table A1.3.10...")
        a1_3_10(db_path)
        progress_dlg.Update(12, "Importing Table A1.3.11...")
        a1_3_11(db_path)
        progress_dlg.Update(13, "Importing Table A1.3.12...")
        a1_3_12(db_path)
        progress_dlg.Update(14, "Importing Table A1.3.13...")
        a1_3_13(db_path)
        progress_dlg.Update(15, "Importing Table A1.3.14...")
        a1_3_14(db_path)
        progress_dlg.Update(16, "Importing Table A1.3.15...")
        a1_3_15(db_path)
        progress_dlg.Update(17, "Importing Table A1.3.16...")
        a1_3_16(db_path)
        progress_dlg.Update(18, "Importing Table A1.3.17...")
        a1_3_17(db_path)
        progress_dlg.Update(19, "Importing Table A1.3.18...")
        a1_3_18(db_path)
        progress_dlg.Update(20)
        wx.MessageBox("Databook tables have been imported successfully.", "Import Complete", wx.OK | wx.ICON_INFORMATION)
    except:
        Visum.Log(20480, traceback.format_exc())
        progress_dlg.Destroy()
        wx.MessageBox("Error while importing data.\nPlease check the Visum log files for more information.", "Error", wx.OK | wx.ICON_ERROR)

