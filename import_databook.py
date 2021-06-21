import os
import sys
sys.path.append(r"C:\Users\david.aspital\Anaconda3\envs\MLSAT\Lib\site-packages")
import wx
import datetime
import pandas as pd
import numpy as np



def create_attributes(db_path):

    # Attributes to be added or updated
    atts = {'DB_IMPORT_DATETIME' : datetime.datetime.now().strftime(r'%d-%m-%Y_%H-%M-%S'),
            'DB_USER' : os.getusername(),
            'DB_PATH' : db_path,
            'DB_VERSION' : pd.read_excel(db_path, sheet_name='Cover', skiprows=2, columns='A')['TAG Data Book'][0]
            }

    # Try to add attribute (ignore if already exists), then update value
    for att, value in atts.items():
        try:
            Visum.Net.AddUserDefinedAttribute(att, att, att, 5)
        except:
            pass
        Visum.Net.SetAttValue(att, value)


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
            if typ == 'int':
                typ = 1
            elif typ == 'float64':
                typ = 2
            elif typ == 'O':
                typ = 5
            else:
                raise ValueError
            udt.TableEntries.AddUserDefinedAttribute(uda_id, col, col, typ)
        udt.TableEntries.SetMultiAttValues(uda_id, tuple(zip(range(1, len(df)+1), df[col].tolist())))


def a1_1_1(db_path):
    name = 'A1.1.1'
    comment = 'Table A 1.1.1:  Green Book Discount Rates'
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
    comment = "Table A 1.3.1: Values of Working (Employers' Business) Time by Mode (£ per hour)"
    df = pd.read_excel(db_path, sheet_name=name, skiprows=23, usecols='A:F', skipfooter=19, header=1, engine='openpyxl')
    df.drop(0, inplace=True)
    df.rename({'Factor':'Factor Cost', 'Perceived':'Perceived Cost', 'Market':'Market Price'}, axis=1, inplace=True)
    df = df[['Mode', 'Factor Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}a', comment)

    comment = 'Values of Non-Working Time by Trip Purpose (£ per hour)' 
    df = pd.read_excel(db_path, sheet_name=name, skiprows=41, usecols='A:F', skipfooter=13, header=1, engine='openpyxl')
    df.drop(0, inplace=True)
    df.rename({'Factor':'Factor Cost', 'Perceived':'Perceived Cost', 'Market':'Market Price'}, axis=1, inplace=True)
    df = df[['Trip Purpose', 'Factor Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}b', comment)

    comment = "Parameter values for employers' business value of time by mode"
    df = pd.read_excel(db_path, sheet_name=name, skiprows=38, skipfooter=12, usecols='H:J', engine='openpyxl')
    create_fill_udt(df, f'{name}c', comment)

    comment = "Values of Working (Employers' Business) Time by mode per person (distance banded)" 
    df = pd.read_excel(db_path, sheet_name=name, skiprows=48, usecols='H:K', header=1, engine='openpyxl')
    df.drop(0, inplace=True)
    df['Distance Band'] = df['Mode'].str.split(" ").str[-1]
    df['Mode'] = df['Mode'].str.split(" ").str[:-1]
    df['Mode'] = [" ".join(map(str, l)) for l in df['Mode']]
    df[['Lower Bound', 'Upper Bound']] = df['Distance Band'].str.split('-', expand=True)
    df['Lower Bound'] = df['Lower Bound'].str.replace('km', '').str.replace('+', '')
    df['Upper Bound'] = df['Upper Bound'].str.replace('km', '')
    df['Upper Bound'].fillna(999999, inplace=True)
    df[['Lower Bound', 'Upper Bound']] = df[['Lower Bound', 'Upper Bound']].astype(int)
    df.rename({'Resource':'Resource Cost', 'Perceived':'Perceived Cost', 'Market':'Market Price'}, axis=1, inplace=True)
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
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, D:Q', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}a', comment)

    comment = 'Forecast values of time per person - Non-Working - Resource cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, R:S', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, R:S', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}b', comment)

    comment = 'Forecast values of time per person - Working - Perceived cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, T:AG', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, T:AG', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}c', comment)

    comment = 'Forecast values of time per person - Non-Working - Perceived cost values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AH:AI', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, AH:AI', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}d', comment)

    comment = 'Forecast values of time per person - Working - Market price values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AJ:AW', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, AJ:AW', header=None, names=header, engine='openpyxl')
    create_fill_udt(df, f'{name}e', comment)

    comment = 'Forecast values of time per person - Non-Working - Market price values (£ per hour)'
    header_cells = pd.read_excel(db_path, sheet_name=name, skiprows=24, usecols='B, AX:AY', nrows=2, header=None, engine='openpyxl').fillna("")
    header = []
    for col in list(header_cells.columns.values):
        header.append(header_cells[col].str.cat(sep = " "))
    header[0] = 'Year'
    header = [x.rstrip() for x in header]
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='B, AX:AY', header=None, names=header, engine='openpyxl')
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
    df = pd.read_excel(db_path, sheet_name=name, skiprows=39, usecols='A,B,H:J', nrows=7, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
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
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='A,B,D:J', nrows=7, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df = df.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Distance Travelled by Vehicles')
    df2 = pd.read_excel(db_path, sheet_name=name, skiprows=34, usecols='A,B,D:J', nrows=12, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df2 = df2.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Distance Travelled by Occupants')
    df3 = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='A,B,K:Q', nrows=7, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df3 = df3.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Vehicle Trips')
    df4 = pd.read_excel(db_path, sheet_name=name, skiprows=34, usecols='A,B,K:Q', nrows=12, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df4 = df4.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Percentage of Person Trips')
    
    df5 = df.merge(df2).merge(df3).merge(df4)
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
    df = pd.read_excel(db_path, sheet_name=name, skiprows=26, usecols='A,B,D:J', nrows=12, header=None, names=header, engine='openpyxl', index_col=[0,1]).reset_index()
    df = df.melt(id_vars=['Mode','Journey Purpose'], value_vars=['7am – 10am', '10am – 4pm', '4pm – 7pm', '7pm – 7am', 'Average Weekday', 'Weekend Average', 'All Week Average'], var_name='Time Period', value_name='Market Price Value of Time Per Vehicle')
    df['Mode'] = np.where(df.Mode == 'PSV ', 'PSV (Occupants)', np.where(df.Mode == '(Occupants)', 'PSV (Occupants)', df.Mode))
    create_fill_udt(df, f'{name}', comment)

def a1_3_6(db_path):
    name = 'A1.3.6'
    comment = 'Market Price Values of Time per Vehicle based on distance travelled (£ per hour)'
    df = pd.read_excel(db_path, sheet_name=name, nrows=80, skiprows=23, header=[0, 1,2,3], engine='openpyxl')
    df.dropna(axis=1, inplace=True)
    df.columns = ['Year' if 'Year' in col else ','.join(col).strip() for col in df.columns.values]
    df = df.melt(id_vars='Year', var_name='Variables', value_name='Market Price Value of Time Per Vehicle')
    df[['Day Type', 'Time Period', 'Mode', 'Journey Purpose']] = df['Variables'].str.split(',', expand=True)
    df = df[['Year', 'Day Type', 'Time Period', 'Mode', 'Journey Purpose', 'Market Price Value of Time Per Vehicle']]
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
    df.columns = ['Vehicle Category' if 'Vehicle' in col else ','.join(col).strip() for col in df.columns.values]
    df2 = pd.read_excel(db_path, sheet_name=name, nrows=4, skiprows=35, header=None,  names=df.columns.values.tolist(), engine='openpyxl', usecols='A,D:I')
    df2.fillna(0, inplace=True)
    df3 = df.append(df2, ignore_index=True)
    #! Change first col to vehicle type, tidy up column names
    create_fill_udt(df3, f'{name}', comment)




if __name__ == '__main__':
    app = wx.App()
    wildcard = "Excel Files(*.xlsm; *.xlsx)|*.xlsm;*.xlsx|" "All files (*.*)|*.*"
    db_path = file_select_dlg("Please select TAG Databook file...", wildcard)
    #a1_1_1(db_path)
    #a_1_3_1(db_path)
    #a1_3_2(db_path)
    #a1_3_3(db_path)
    #a1_3_4(db_path)
    #a1_3_5(db_path)
    #a1_3_6(db_path)
    #a1_3_7(db_path)
    a1_3_8(db_path)

    print(db_path)