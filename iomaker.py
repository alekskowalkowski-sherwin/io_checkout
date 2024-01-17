import pandas as pd
import numpy as np
import openpyxl



pd.set_option('display.max_columns', None)

valve_descriptors = {
    "replace_string": "VALVE",
    "Closed Limit Switch": "I_VALVE_XSC",
    "Opened Limit Switch": "I_VALVE_XSO",
    "Solenoid": "O_VALVE_XS",
    "Solenoid 1": "O_VALVE_XS1",
    "Solenoid 2": "O_VALVE_XS2"
}

vpr_descriptors = {
    "replace_string": "VALVE",
    "Flow Output": "O_VALVE_FV",
    "Temperature Output": "O_VALVE_TV",
    "Pressure Output": "O_VALVE_PV"
}

def find_io_point(io_type, io_point_name):
    temp_dict = {}
    io_point = tags_df.loc[tags_df["NAME"] == io_point_name]["SPECIFIER"].values[0] if not tags_df.loc[tags_df["NAME"] == io_point_name].empty else None
    if io_point is not None:
        temp_dict[io_type] = io_point
    else:
        temp_dict[io_type] = "NA"
    return temp_dict

def generate_df(descriptors, device_list):
    device_dict = {}
    for device in device_list:
        temp_dict = {}
        for descriptor in descriptors.keys():
            if descriptor == "replace_string":
                continue
            io_name = descriptors[descriptor].replace(descriptors["replace_string"], device)
            temp_dict = {**temp_dict, **find_io_point(descriptor, io_name)}
        device_dict[device] = temp_dict
        device_df = pd.DataFrame.from_dict(device_dict, orient="index").reset_index().rename(columns={"index": "Device"})
    return device_df


if __name__ == "__main__":

    valve_dict = {}
    vpr_dict = {}

    tags_df = pd.read_excel("./tags.xlsx")
    reference_vpr_df = pd.read_excel("./1A - Valve Automated Checkout.xlsx", sheet_name="New Control Valve List")
    reference_valve_df = pd.read_excel("./1A - Valve Automated Checkout.xlsx", sheet_name="New Actuated Valve List")

    vpr_list = reference_vpr_df["Valve Number"].tolist()
    valve_list = reference_valve_df["Valve Number"].tolist()

    with pd.ExcelWriter('test.xlsx', engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        generate_df(valve_descriptors, valve_list).to_excel(writer, "Valves")
        generate_df(vpr_descriptors, vpr_list).to_excel(writer, "VPR")

