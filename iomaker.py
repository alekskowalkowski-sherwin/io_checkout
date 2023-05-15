import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import re
from pathlib import Path
import os
def xlookup(lookup_value, lookup_array, return_array, if_not_found: str = ''):
    match_value = return_array.loc[lookup_array == lookup_value]
    if match_value.empty:
        return f'not found' if if_not_found == '' else if_not_found

    else:
        return match_value.tolist()[0]


#Get file path
path_tags=os.getcwd()+'/tags.xlsx'
path_checkout=os.getcwd()+'/1A - Valve Automated Checkout.xlsx'

#Set Dataframes in File Path, for the tag export and checkout sheet
df_tags=pd.read_excel(path_tags)
df_checkout=pd.read_excel(path_checkout, sheet_name="New Actuated Valve List")
df_checkout=df_checkout.fillna("")

df_vpr=pd.read_excel(path_checkout, sheet_name="New Control Valve List")
df_vpr=df_vpr.fillna("")

#Define holder lists to append stuff to
close_list=[]
open_list=[]
sol_list=[]
valve_list=[]

flowcr_list=[]
tempcr_list=[]
prescr_list=[]

def v2w_lookup(row):
    valve = row['Valve Number']
    close_switch_id=f'I_{valve}_XSC'
    open_switch_id=f'I_{valve}_XSO'
    solenoid_id=f'O_{valve}_XS'
    solenoid1_id=f'O_{valve}_XS1'
    solenoid2_id=f'O_{valve}_XS2'

    close_switch_address= df_tags.loc[df_tags['NAME'] == close_switch_id, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == close_switch_id, 'SPECIFIER'].empty else None
    open_switch_address= df_tags.loc[df_tags['NAME'] == open_switch_id, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == open_switch_id, 'SPECIFIER'].empty else None
    solenoid_address= df_tags.loc[df_tags['NAME'] == solenoid_id, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == solenoid_id, 'SPECIFIER'].empty else None
    solenoid1_address= df_tags.loc[df_tags['NAME'] == solenoid1_id, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == solenoid1_id, 'SPECIFIER'].empty else None
    solenoid2_address= df_tags.loc[df_tags['NAME'] == solenoid2_id, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == solenoid2_id, 'SPECIFIER'].empty else None

    close_list.append(close_switch_address)
    open_list.append(open_switch_address)
    sol_list.append(solenoid_address)
    valve_list.append(valve)

    print(f'{valve} IO: XSC-{close_switch_address}, XSO-{open_switch_address}, XS-{solenoid_address}, {solenoid1_address}, {solenoid2_address})')


def vpr_lookup(row):
    valve = row['Valve Number']
    flow_output=f'O_{valve}_FV'
    temp_output=f'O_{valve}_TV'
    pressure_output=f'O_{valve}_PV'
    
    flow_address= df_tags.loc[df_tags['NAME'] == flow_output, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == flow_output, 'SPECIFIER'].empty else None
    temp_address= df_tags.loc[df_tags['NAME'] == temp_output, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == temp_output, 'SPECIFIER'].empty else None
    pressure_address= df_tags.loc[df_tags['NAME'] == pressure_output, 'SPECIFIER'].iloc[0] if not df_tags.loc[df_tags['NAME'] == pressure_output, 'SPECIFIER'].empty else None
    
    flowcr_list.append(flow_address)
    tempcr_list.append(temp_address)
    prescr_list.append(pressure_address)
    
    print(f'{valve} IO: FV-{flow_output}, TV-{temp_output}, PV-{pressure_output}')
    
          
    
for rows, row in df_checkout.iterrows():
    addresses_v2w=v2w_lookup(row)
for rows, row in df_vpr.iterrows():
    addresses_vpr=vpr_lookup(row)
    


