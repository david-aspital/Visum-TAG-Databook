# Import required libraries
import os
import sys
sys.path.append(f"{os.path.dirname(os.path.realpath(__file__))}\\src")
import wx
import datetime
import pandas as pd
import numpy as np
import traceback

# Create initial network UDAs
def create_attributes():

    # Create / overwrite default modelled year (can be overwritten in visum later)
    model_year = 2021
    model_tp = 'Average Weekday'
    no_of_iter_for_conv = 2
    OGV1_proportion = 0.4
    OGV2_proportion = 0.6
    HGV_VOT_factor = 2.5
    override_avg_net_speed = True

    # Attributes to be added or updated and their types
    atts = {'MODEL_YEAR' : (1, model_year),
            'MODEL_TP' : (5, model_tp),
            'NO_OF_ITER_FOR_CONV' : (1, no_of_iter_for_conv),
            'OGV1_PROPORTION' : (2, OGV1_proportion),
            'OGV2_PROPORTION' : (2, OGV2_proportion),
            'HGV_VOT_FACTOR' : (2, HGV_VOT_factor),
            'OVERRIDE_AVG_NET_SPEED' : (9, override_avg_net_speed)}

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

# Lookup Perceived Value of Time
def Perceived_VOT_int():
    name = 'Perceived_VOT_int'
    comment = 'Interim Perceived Value of Time - Goods Vehicle Disaggregated'
    
    # Try to set table to udt if it exists, otherwise add new table definition
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
    except:
        udt = Visum.Net.AddTableDefinition(name)
    
    # Provide description for the table and define demand seg / vehicle type / journey purpose combinations
    udt.SetAttValue('Comment', comment)
    VT_JP = [['CB', 'Car', 'Work'],
             ['CC', 'Car', 'Commuting'],
             ['CO', 'Car', 'Other'],
             ['LGV', 'LGV', 'Work (freight)'],
             ['LGV', 'LGV', 'Commuting & Other'],
             ['HGV', 'OGV1', 'Working'],
             ['HGV', 'OGV2', 'Working']]
    
    # Define time period strings as they appear in the WebTAG databook tables
    TP = ['7am – 10am','10am – 4pm','4pm – 7pm','7pm – 7am','Average Weekday','Weekend Average','All Week Average']
    
    # Create all possible combinations of VT_JP and TP and add blank rows to table of equal numberthem as table entries
    populate = [[vt_jp[0], vt_jp[1],vt_jp[2], tp] for tp in TP for vt_jp in VT_JP]
    udt.AddMultiTableEntries(range(1,len(populate)+1))
    
    # Define column names and iterate through columns, deleting and re-adding if they already exist
    IDnm = ['AUC', 'Vehicle_Type', 'Journey_Purpose', 'Time_Period']
    for i, id in enumerate(IDnm):
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 5)
        
        # Set values of blank rows to the values in relevant column of populate
        udt.TableEntries.SetMultiAttValues(id, tuple(zip(range(1, len(populate)+1), [element[i] for element in populate])))
    
    # Define a new column name and create a string in the Visum formula language to perform a lookup
    IDnm = 'Value_of_Time_Per_Vehicle'
    Condition = '(A[Mode]=[Vehicle_Type])&(A[Journey_Purpose]=[Journey_Purpose])&(A[Time_Period]=[Time_Period])'
    
    # Define a string in Visum formula language to return the reduction factor to apply for indirect tax correction
    ITCD = f'If([Journey_Purpose]=\"Work\", [NETWORK\INDIRECT_TAX_CORRECTION], 1)'
    
    # Delete and re-add the column if it already exists, then create a formula attribute to lookup from WebTAG table A1.3.5
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'TableLookup(TABLEENTRIES_A1_3_5 A, {Condition}, A[Market_Price_{IDnm}]/{ITCD})')

# Weighted Averages by DSeg for Perceived Value of Time
def Perceived_VOT_final():
    name = 'Perceived_VOT_final'
    comment = 'Final Perceived Value of Time - Goods Vehicle Aggregated'
    
    # Try to set table to udt if it exists, otherwise add new table definition
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
    except:
        udt = Visum.Net.AddTableDefinition(name)
    
    # Provide description for the table and define demand seg (with vehicle type / journey purpose descriptors aggregated appropriately)
    udt.SetAttValue('Comment', comment)
    VT_JP = [['CB', 'Car', 'Work'],
             ['CC', 'Car', 'Commuting'],
             ['CO', 'Car', 'Other'],
             ['LGV', 'LGV', 'Average LGV'],
             ['HGV', 'HGV', 'Working']]
    
    # Define time period strings as they appear in the WebTAG databook tables
    TP = ['7am - 10am','10am - 4pm','4pm - 7pm','7pm - 7am','Average Weekday','Weekend Average','All Week Average']
    
    # Create all possible combinations of VT_JP and TP and add blank rows to table of equal numberthem as table entries
    populate = [[vt_jp[0],vt_jp[1],vt_jp[2], tp] for tp in TP for vt_jp in VT_JP]
    udt.AddMultiTableEntries(range(1,len(populate)+1))
    
    # Define column names and iterate through columns, deleting and re-adding if they already exist
    IDnm = ['AUC', 'Vehicle_Type', 'Journey_Purpose', 'Time_Period']
    for i, id in enumerate(IDnm):
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 5)
        
        # Set values of blank rows to the values in relevant column of populate
        udt.TableEntries.SetMultiAttValues(id, tuple(zip(range(1, len(populate)+1), [element[i] for element in populate])))
    
    # Define column name for formula attribute
    IDnm = 'Value_of_Time_Per_Vehicle'
    
    #Look up the relative proportions of LGVs using Visum formula language test strings (string values of A1_3_4_work/non_work defined later) and apply them
    LGV_work = f'{A1_3_4_work}*TableLookup(TABLEENTRIES_Perceived_VOT_int A, (A[AUC]=\"LGV\")&(A[Time_Period]=[Time_Period])&(A[Journey_Purpose]=\"Work (freight)\"), A[{IDnm}])/100'
    LGV_non_work = f'{A1_3_4_non_work}*TableLookup(TABLEENTRIES_Perceived_VOT_int A, (A[AUC]=\"LGV\")&(A[Time_Period]=[Time_Period])&(A[Journey_Purpose]=\"Commuting & Other\"), A[{IDnm}])/100'
    
    #Return relative proportions of HGVs using assumed valued stored in network UDAs created above and apply them
    OGV1 = f'[NETWORK\OGV1_PROPORTION]*TableLookup(TABLEENTRIES_Perceived_VOT_int A, (A[Vehicle_Type]=\"OGV1\")&(A[Time_Period]=[Time_Period]), A[{IDnm}])'
    OGV2 = f'[NETWORK\OGV2_PROPORTION]*TableLookup(TABLEENTRIES_Perceived_VOT_int A, (A[Vehicle_Type]=\"OGV2\")&(A[Time_Period]=[Time_Period]), A[{IDnm}])'
    
    #Return the values for car by journey purpose as they were in the interim table, without any weighted averages being required
    not_GV = f'TableLookup(TABLEENTRIES_Perceived_VOT_int A, (A[AUC]=[AUC])&(A[Time_Period]=[Time_Period]), A[{IDnm}])'
    
    #Use the above Visum formula language strings to construct the new formula UDA
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'If([AUC]=\"LGV\", {LGV_work}+{LGV_non_work}, If([AUC]=\"HGV\", [NETWORK\HGV_VOT_FACTOR]*({OGV1}+{OGV2}), {not_GV}))')
    
    # Create a new formula UDA
    IDnm = 'VOT_pence_per_sec'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = '[Value_of_Time_Per_Vehicle]/36')

def Perceived_VOC_int():
    name = 'Perceived_VOC_int'
    comment = 'PLACEHOLDER'
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
    except:
        udt = Visum.Net.AddTableDefinition(name)
    udt.SetAttValue('Comment', comment)
    VT_JP = [['CB', 'Car', 'Work', 'All', 65],
             ['CC', 'Car', 'Non-Work', 'All', 54],
             ['CO', 'Car', 'Non-Work', 'All', 54],
             ['LGV', 'LGV', 'Work', 'Non-electric', 54],
             ['LGV', 'LGV', 'Work', 'Electric', 54],
             ['LGV', 'LGV', 'Non-Work', 'Non-electric', 54],
             ['LGV', 'LGV', 'Non-Work', 'Electric', 54],
             ['HGV', 'OGV1', 'Work', 'All', 65],
             ['HGV', 'OGV2', 'Work', 'All', 65]]
    udt.AddMultiTableEntries(range(1,len(VT_JP)+1))
    IDnm = ['AUC', 'Vehicle_Type', 'Trip_Purpose', 'Fuel_Type', 'Override_Avg_Net_Speed_kph']
    IDnmType = [5, 5, 5, 5, 2]
    for i, id in enumerate(IDnm):
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, IDnmType[i])
        udt.TableEntries.SetMultiAttValues(id, tuple(zip(range(1, len(VT_JP)+1), [element[i] for element in VT_JP])))
    IDnm = 'Used_Avg_Speed'
    DSegSetCode = 'TableLookup(DEMANDSEGMENT D, (V[DSegSetCode]=D[Code]), D[MODE\TSYSSET])'
    VehKmTravPrT = f'TableLookup(PRTASSIGNMENTQUALITY V, (V[Iteration]=[Network\\NO_OF_ITER_FOR_CONV])&({DSegSetCode}=[AUC]), V[VehKmTravPrT])'
    VehHourTravtCur = f'TableLookup(PRTASSIGNMENTQUALITY V,(V[Iteration]=[Network\\NO_OF_ITER_FOR_CONV])&({DSegSetCode}=[AUC]), V[VehHourTravtCur])'
    Calc_Avg_Net_Speed_kph = f'{VehKmTravPrT}/{VehHourTravtCur}' #CHECK UNITS
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'If([Network\OVERRIDE_AVG_NET_SPEED], [Override_Avg_Net_Speed_kph], {Calc_Avg_Net_Speed_kph})')
    IDnm = ['Param_a', 'Param_b', 'Param_c', 'Param_d']
    for i, id in enumerate(IDnm):
        A1_3_12 = f'TableLookup(TABLEENTRIES_A1_3_12 A,(A[Year]=[NETWORK\MODEL_YEAR])&(A[Vehicle_Type]=[Vehicle_Type])&(A[Fuel_Type]=\"Average\"), A[{id}])'
        A1_3_13 = f'TableLookup(TABLEENTRIES_A1_3_13 A,(A[Year]=[NETWORK\MODEL_YEAR])&(A[Vehicle_Type]=[Vehicle_Type]), A[{id}])'
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 2, formula = f'If([Trip_Purpose]=\"Work\", {A1_3_12}, {A1_3_13})')
        ModeLookup = f'IF([CODE]="CB"|[CODE]="CC"|[CODE]="CO",TableLookup(TABLEENTRIES_PERCEIVED_VOC_INT A, A[AUC]=[CODE], A[{id}]),0/0)'
        try:
            Visum.Net.Modes.AddUserDefinedAttribute(id, id, id, 2, formula = ModeLookup)
        except:
            pass
    IDnm = 'VOC_f'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'[PARAM_A]/[Used_Avg_Speed]+[PARAM_B]+[PARAM_C]*[Used_Avg_Speed]+[PARAM_D]*POW([Used_Avg_Speed],2)')
    IDnm = ['Param_a1', 'Param_b1']
    for i, id in enumerate(IDnm):
        A1_3_14 = f'TableLookup(TABLEENTRIES_A1_3_14 A,((A[Vehicle_Type]=[Vehicle_Type])&(A[Trip_Purpose]=[Trip_Purpose])&A[Fuel_Type]=[Fuel_Type]), A[{id}])'
        A1_3_15 = f'TableLookup(TABLEENTRIES_A1_3_15 A,((A[Vehicle_Type]=[Vehicle_Type])&(A[Trip_Purpose]=[Trip_Purpose])&A[Year]=[NETWORK\MODEL_YEAR]), A[{id}])'
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 2, formula = f'If([Trip_Purpose]=\"Non-Work\", 0, If([Vehicle_Type]=\"Car\", {A1_3_15}, {A1_3_14}))')
        ModeLookup = f'IF([CODE]="CB"|[CODE]="CC"|[CODE]="CO",TableLookup(TABLEENTRIES_PERCEIVED_VOC_INT A, A[AUC]=[CODE], A[{id}]),0/0)'
        try:
            Visum.Net.Modes.DeleteUserDefinedAttribute(id)
        except:
            pass
        Visum.Net.Modes.AddUserDefinedAttribute(id, id, id, 2, formula = ModeLookup)
    IDnm = 'VOC_nf'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'[PARAM_A1]+[PARAM_B1]/[Used_Avg_Speed]')

def Perceived_VOC_final():
    name = 'Perceived_VOC_final'
    comment = 'PLACEHOLDER'
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
    except:
        udt = Visum.Net.AddTableDefinition(name)
    udt.SetAttValue('Comment', comment)
    VT_JP = [['CB', 'Car', 'Work'],
             ['CC', 'Car', 'Non-Work'],
             ['CO', 'Car', 'Non-Work'],
             ['LGV', 'LGV', 'Average LGV'],
             ['HGV', 'HGV', 'Work']]
    udt.AddMultiTableEntries(range(1,len(VT_JP)+1))
    IDnm = ['AUC', 'Vehicle_Type', 'Trip_Purpose']
    for i, id in enumerate(IDnm):
        try:
            udt = Visum.Net.TableDefinitions.ItemByKey(name)
        except:
            udt = Visum.Net.AddTableDefinition(name)
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 5)
        udt.TableEntries.SetMultiAttValues(id, tuple(zip(range(1, len(VT_JP)+1), [element[i] for element in VT_JP])))
    A1_3_9_e = f'TableLookup(TABLEENTRIES_A1_3_9 A,(A[Year]=[NETWORK\MODEL_YEAR])&(A[Mode]=\"LGV\")&(A[Fuel_Type]=\"Electric\"), A[Value])'
    IDnm = ['VOC_f', 'VOC_nf']
    for i, id in enumerate(IDnm):
        LGV_work_non_e = f'(1-{A1_3_9_e})*TableLookup(TABLEENTRIES_Perceived_VOC_int A, (A[AUC]=\"LGV\")&(A[Trip_Purpose]=\"Work\")&(A[Fuel_Type]=\"Non-electric\"), A[{id}])'
        LGV_work_e = f'{A1_3_9_e}*TableLookup(TABLEENTRIES_Perceived_VOC_int A, (A[AUC]=\"LGV\")&(A[Trip_Purpose]=\"Work\")&(A[Fuel_Type]=\"Electric\"), A[{id}])'
        LGV_non_work_non_e = f'(1-{A1_3_9_e})*TableLookup(TABLEENTRIES_Perceived_VOC_int A, (A[AUC]=\"LGV\")&(A[Trip_Purpose]=\"Non-Work\")&(A[Fuel_Type]=\"Non-electric\"), A[{id}])'
        LGV_non_work_e = f'{A1_3_9_e}*TableLookup(TABLEENTRIES_Perceived_VOC_int A, (A[AUC]=\"LGV\")&(A[Trip_Purpose]=\"Non-Work\")&(A[Fuel_Type]=\"Electric\"), A[{id}])'
        LGV = f'({A1_3_4_work}*({LGV_work_non_e}+{LGV_work_e})+{A1_3_4_non_work}*({LGV_non_work_non_e}+{LGV_non_work_e}))/100'
        OGV1 = f'[NETWORK\OGV1_PROPORTION]*TableLookup(TABLEENTRIES_Perceived_VOC_int A, A[Vehicle_Type]=\"OGV1\", A[{id}])'
        OGV2 = f'[NETWORK\OGV2_PROPORTION]*TableLookup(TABLEENTRIES_Perceived_VOC_int A, A[Vehicle_Type]=\"OGV2\", A[{id}])'
        not_GV = f'TableLookup(TABLEENTRIES_Perceived_VOC_int A, A[AUC]=[AUC], A[{id}])'
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 2, formula = f'If([AUC]=\"LGV\", {LGV}, If([AUC]=\"HGV\", {OGV1}+{OGV2}, {not_GV}))')
    IDnm = 'VOC'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = '[VOC_f]+[VOC_nf]')
    IDnm = 'VOC_pence_per_m'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = '[VOC]/1000')
    
def UDAs_for_Impedance():
    name = 'UDAs_for_Impedance'
    comment = 'PLACEHOLDER'
    try:
        udt = Visum.Net.TableDefinitions.ItemByKey(name)
    except:
        udt = Visum.Net.AddTableDefinition(name)
    udt.SetAttValue('Comment', comment)
    AUC = ['CB', 'CC', 'CO', 'LGV', 'HGV']
    TERM = ['DIST', 'TIME', 'TOLL']
    populate = [[auc, term] for term in TERM for auc in AUC]
    udt.AddMultiTableEntries(range(1,len(populate)+1))
    IDnm = ['AUC', 'TERM']
    for i, id in enumerate(IDnm):
        try:
            udt.TableEntries.DeleteUserDefinedAttribute(id)
        except:
            pass
        udt.TableEntries.AddUserDefinedAttribute(id, id, id, 5)
        udt.TableEntries.SetMultiAttValues(id, tuple(zip(range(1, len(populate)+1), [element[i] for element in populate])))
    IDnm = 'Value'
    DIST = 'TableLookup(TABLEENTRIES_Perceived_VOC_final A, A[AUC]=[AUC], A[VOC_pence_per_m])/TableLookup(TABLEENTRIES_Perceived_VOT_final A, (A[AUC]=[AUC])&(A[Time_Period]=[NETWORK\MODEL_TP]), A[VOT_pence_per_sec])'
    TIME = '1'
    TOLL = '1/TableLookup(TABLEENTRIES_Perceived_VOT_final A, (A[AUC]=[AUC])&(A[Time_Period]=[NETWORK\MODEL_TP]), A[VOT_pence_per_sec])'
    try:
        udt.TableEntries.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    udt.TableEntries.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'If([TERM]=\"DIST\", {DIST}, If([TERM]=\"TIME\", {TIME}, {TOLL}))')
    for i, id in enumerate(populate):
        AUC = id[0]
        TERM = id[1]
        IDnm = f'{AUC}_IMP_{TERM}'
        try:
            Visum.Net.DeleteUserDefinedAttribute(IDnm)
        except:
            pass
        Visum.Net.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = f'TableLookup(TABLEENTRIES_UDAs_for_Impedance A, (A[AUC]=\"{AUC}\")&(A[TERM]=\"{TERM}\"), A[Value])')

def Activity_Pair_UDAs():
    IDnm = 'OCC'
    OCC = 'TableLookup(TABLEENTRIES_A1_3_3A A, (A[AUC]=[AUC])&(A[Time_Period]=[NETWORK\MODEL_TP]), A[Occupancy_Per_Trip])'
    try:
        Visum.Net.ActPairs.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    Visum.Net.ActPairs.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = OCC)
    CB_VoT = 'TableLookup(TABLEENTRIES_A1_3_2C A, (A[Year]=[NETWORK\MODEL_YEAR]), A[Car_driver])'
    CC_VoT = 'TableLookup(TABLEENTRIES_A1_3_2D A, (A[Year]=[NETWORK\MODEL_YEAR]), A[Commuting])'
    CO_VoT = 'TableLookup(TABLEENTRIES_A1_3_2D A, (A[Year]=[NETWORK\MODEL_YEAR]), A[Other])'
    IDnm = 'VOT'
    VOT = f'IF([AUC]=\"CB\", {CB_VoT},IF([AUC]=\"CC\", {CC_VoT}, {CO_VoT}))'
    try:
        Visum.Net.ActPairs.DeleteUserDefinedAttribute(IDnm)
    except:
        pass
    Visum.Net.ActPairs.AddUserDefinedAttribute(IDnm, IDnm, IDnm, 2, formula = VOT)

A1_3_4_work = f'TableLookup(TABLEENTRIES_A1_3_4 A, (A[Journey_Purpose]=\"Work (freight)\")&(A[Mode]=\"LGV\")&(A[Time_Period]=[NETWORK\MODEL_TP]), A[Percentage_of_Vehicle_Trips])'
A1_3_4_non_work = f'TableLookup(TABLEENTRIES_A1_3_4 A, (A[Journey_Purpose]=\"Non - Work\")&(A[Mode]=\"LGV\")&(A[Time_Period]=[NETWORK\MODEL_TP]), A[Percentage_of_Vehicle_Trips])'


if __name__ == '__main__':
    app = wx.App()
    num_tables = 6

    try:
        progress_dlg = wx.ProgressDialog("Importing Tables", "Importing tables from databook...", num_tables+1, style=wx.PD_APP_MODAL | wx.PD_SMOOTH | wx.PD_AUTO_HIDE)
        create_attributes()
        progress_dlg.Update(1, "Creating Table Perceived_VOT_int...")
        Perceived_VOT_int()
        progress_dlg.Update(2, "Creating Table Perceived_VOT_final...")
        Perceived_VOT_final()
        progress_dlg.Update(3, "Creating Table Perceived_VOC_int...")
        Perceived_VOC_int()
        progress_dlg.Update(4, "Creating Table Perceived_VOC_final...")
        Perceived_VOC_final()
        progress_dlg.Update(5, "Creating Table UDAs_for_Impedance...")
        UDAs_for_Impedance()
        progress_dlg.Update(6, "Applying Activity_Pair_UDAs...")
        Activity_Pair_UDAs()
        progress_dlg.Update(7)
        wx.MessageBox("All processing tables have been populated successfully.", "Processing Complete", wx.OK | wx.ICON_INFORMATION)
    except:
        Visum.Log(20480, traceback.format_exc())
        progress_dlg.Destroy()
        wx.MessageBox("Error while processing data.\nPlease check the Visum log files for more information.", "Error", wx.OK | wx.ICON_ERROR)