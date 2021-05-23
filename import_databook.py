import os
import sys
import wx
import datetime
import pandas as pd



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
    df = pd.read_excel(db_path, sheet_name=name, skiprows=23)
    df['Years from current year'] = df['Unnamed: 1']
    df = df[['Years from current year', 'Discount rate']]
    df[['Lower Bound', 'Upper Bound']] = df['Years from current year'].str.split('-', expand=True)
    df['Upper Bound'].fillna(99999, inplace=True)
    df['Lower Bound'] = df['Lower Bound'].str.replace(' and over', '')
    df[['Lower Bound', 'Upper Bound']] = df[['Lower Bound', 'Upper Bound']].astype(int)
    df = df[['Years from current year', 'Lower Bound', 'Upper Bound', 'Discount rate']]
    create_fill_udt(df, name, comment)

def a_1_3_1(db_path):
    name = 'A1.3.1'
    comment = "Table A 1.3.1: Values of Working (Employers' Business) Time by Mode (£ per hour)"
    df = pd.read_excel(db_path, sheet_name=name, skiprows=23, usecols='A:F', skipfooter=6, header=[0,1])
    df.columns = [f'{i} {j}' if j != '' else f'{i}' for i,j in df.columns]
    df = df[['Mode', 'Factor Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}a', comment)

    comment = 'Values of Non-Working Time by Trip Purpose (£ per hour)' 
    df = pd.read_excel(db_path, sheet_name=name, skiprows=41, usecols='A:F', header=[0,1])
    df.columns = [f'{i} {j}' if j != '' else f'{i}' for i,j in df.columns]
    df = df[['Trip Purpose', 'Factor Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}b', comment)

    comment = "Parameter values for employers' business value of time by mode"
    df = pd.read_excel(db_path, sheet_name=name, skiprows=37, usecols='H:J')
    df.columns = [f'{i} {j}' if j != '' else f'{i}' for i,j in df.columns]
    create_fill_udt(df, f'{name}c', comment)

    comment = "Values of Working (Employers' Business) Time by mode per person (distance banded)" 
    df = pd.read_excel(db_path, sheet_name=name, skiprows=48, usecols='H:K', header=[0,1])
    df.columns = [f'{i} {j}' if j != '' else f'{i}' for i,j in df.columns]
    df = df[['Mode', 'Resource Cost', 'Perceived Cost', 'Market Price']]
    create_fill_udt(df, f'{name}d', comment)

def a1_3_2(db_path):
    name = 'A1.3.2'
    comment = 'Forecast values of time per person - Resource cost values (£ per hour)'
    df = pd.read_excel(db_path, sheet_name=name, skiprows=23, usecols='A:S', header=[0,1])





if __name__ == '__main__':
    app = wx.App()
    wildcard = "Excel Files(*.xlsm; *.xlsx)|*.xlsm;*.xlsx|" \
         "All files (*.*)|*.*"
    db_path = file_select_dlg("Please select TAG Databook file...", wildcard)
    a1_1_1(db_path)

    print(db_path)