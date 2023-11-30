""" Generates the information needed for the SAR Report's Section 1 """

import pandas as pd
import logging
import os

from pathlib import Path

class Sec_1:
    
    def __init__(self, data: list, transmitter_names: list, sec_1_fcc_filepath: str, sec_1_ised_filepath: str, log_dir: str):
        self.data = pd.concat(data)
        self.trans = transmitter_names
        self.s1_fcc = sec_1_fcc_filepath
        self.s1_ised = sec_1_ised_filepath
        
        self.tp = {
            "Head"           : ["Left Cheek", "Left Tilt", "Right Cheek", "Right Tilt"],
            "Body-worn"      : ["Back", "Front"],
            "Body & Hotspot" : ["Back", "Front", "Edge Top", "Edge Right", "Edge Bottom", "Edge Left"],
            "Hotspot"        : ["Edge Top", "Edge Right", "Edge Bottom", "Edge Left"],
            "Extremity"      : ["Back", "Front", "Edge Top", "Edge Right", "Edge Bottom", "Edge Left"]
            }
        self.tl = {
            "GSM"       :  ["GSM 850", "GSM E-900", "GSM R-900", "GSM 1800", "GSM 1900"],
            "W-CDMA"    :  ["W-CDMA B1", "W-CDMA B2", "W-CDMA B4", "W-CDMA B5", "W-CDMA B8"],
            "LTE"       :  ["LTE B1", "LTE B2", "LTE B3", "LTE B4", "LTE B5", "LTE B7", "LTE B8",
                            "LTE B11", "LTE B12", "LTE B13", "LTE B14", "LTE B17", "LTE B18", "LTE B19",
                            "LTE B20", "LTE B21", "LTE B22", "LTE B24", "LTE B25", "LTE B26", "LTE B27",
                            "LTE B28", "LTE B30", "LTE B31", "LTE B33", "LTE B34", "LTE B35", "LTE B36",
                            "LTE B37", "LTE B38", "LTE B39", "LTE B40 (Block A)", "LTE B40 (Block B)",
                            "LTE B40", "LTE B41 PC3", "LTE B41 PC2","LTE B42", "LTE B43", "LTE B44",
                            "LTE B45", "LTE B46 NAR", "LTE B46 CE", "LTE B47","LTE B48", "LTE B49",
                            "LTE B50", "LTE B51", "LTE B52", "LTE B53", "LTE B65", "LTE B66","LTE B68",
                            "LTE B70", "LTE B71", "LTE B72", "LTE B73", "LTE B74", "LTE B85", "LTE B87", "LTE B88"],
            "FR1"       :  ["FR1 n1", "FR1 n2", "FR1 n3", "FR1 n5", "FR1 n7", "FR1 n8", "FR1 n12", "FR1 n13",
                            "FR1 n14", "FR1 n18","FR1 n20", "FR1 n24", "FR1 n25", "FR1 n26", "FR1 n28", "FR1 n30",
                            "FR1 n34", "FR1 n38", "FR1 n40", "FR1 n40 (Block A)", "FR1 n40 (Block B)", "FR1 n40 CE",
                            "FR1 n41 PC3", "FR1 n41 PC2", "FR1 n41 CE PC3", "FR1 n41 CE PC2", "FR1 n46", "FR1 n46 CE",
                            "FR1 n47", "FR1 n48", "FR1 n48 CE", "FR1 n50", "FR1 n51", "FR1 n53","FR1 n65", "FR1 n66",
                            "FR1 n70", "FR1 n71", "FR1 n74", "FR1 n77 PC3", "FR1 n77 CE PC3", "FR1 n77 (Block A) PC3",
                            "FR1 n77 (Block B) PC3", "FR1 n77 CE (Block B) PC3", "FR1 n77 (Block C) PC3",
                            "FR1 n77 CE (Block C) PC3", "FR1 n77 PC2", "FR1 n77 CE PC2", "FR1 n77 (Block A) PC2",
                            "FR1 n77 (Block B) PC2", "FR1 n77 CE (Block B) PC2", "FR1 n77 (Block C) PC2",
                            "FR1 n77 CE (Block C) PC2", "FR1 n78 PC3", "FR1 n78 IC PC3", "FR1 n78 CE PC3",
                            "FR1 n78 PC2", "FR1 n78 IC PC2", "FR1 n78 CE PC2", "FR1 n79", "FR1 n79 CE", "FR1 n85",
                            "FR1 n90", "FR1 n90 CE", "FR1 n91", "FR1 n92","FR1 n93", "FR1 n94", "FR1 n96", "FR1 n96 CE",
                            "FR1 n100", "FR1 n101", "FR1 n102", "FR1 n102 CE", "FR1 n104", "FR1 n104 CE", "FR1 n255", "FR1 n256"],
            "WLAN"      :  ["Wi-Fi 2.4 GHz", "Wi-Fi 5.2 GHz", "Wi-Fi 5.3 GHz", "Wi-Fi 5.5 GHz", "Wi-Fi 5.8 GHz",
                            "Wi-Fi 5.9 GHz", "Wi-Fi 6E", "U-NII 5", "U-NII 6", "U-NII 7", "U-NII 8"],
            "Bluetooth" :  ["Bluetooth (2.4 GHz)", "Bluetooth (NB U-NII 1)", "Bluetooth (NB U-NII 3)"],
            "Thread"    :  ["Thread (2.4 GHz)", "802.15.4ab"],
            "MSS"       :  ["MSS"]
            }
        self.fcc = {
            "TNE" :    ["LTE B53", "FR1 n53", "MSS"],
            "CBE" :    ["LTE B48", "FR1 n48"],
            "PCE" :    ["GSM 850", "GSM 1900", "W-CDMA B2", "W-CDMA B4", "W-CDMA B5", "LTE B2", "LTE B4",
                        "LTE B5", "LTE B7", "LTE B12", "LTE B13", "LTE B14", "LTE B17", "LTE B25", "LTE B26",
                        "LTE B30", "LTE B38", "LTE B40 (Block A)", "LTE B40 (Block B)", "LTE B41 PC3",
                        "LTE B41 PC2", "LTE B66", "LTE B70", "LTE B71", "LTE B74", "LTE B85", "FR1 n2",
                        "FR1 n5", "FR1 n7", "FR1 n12", "FR1 n13", "FR1 n14", "FR1 n25", "FR1 n26", "FR1 n30",
                        "FR1 n38", "FR1 n40 (Block A)", "FR1 n40 (Block B)", "FR1 n41 PC3", "FR1 n41 PC2", "FR1 n66",
                        "FR1 n70", "FR1 n77 PC3", "FR1 n77 (Block A) PC3", "FR1 n77 (Block B) PC3", "FR1 n77 (Block C) PC3",
                        "FR1 n77 PC2", "FR1 n77 (Block A) PC2", "FR1 n77 (Block B) PC2", "FR1 n77 (Block C) PC2", "FR1 n79 (Narrow) PC3"],
            "DTS" :    ["Wi-Fi 2.4 GHz", "Thread (2.4 GHz)"],
            "NII" :    ["Wi-Fi 5.2 GHz", "Wi-Fi 5.3 GHz", "Wi-Fi 5.5 GHz", "Wi-Fi 5.8 GHz", "Wi-Fi 6E", "U-NII 5",
                        "U-NII 6", "U-NII 7", "U-NII 8", "Bluetooth (NB U-NII 1)", "Bluetooth (NB U-NII 3)", "802.15.4ab"],
            "DSS" :    ["Bluetooth (2.4 GHz)"]
            }
        self.ised = {
            "PCE" :    ["GSM 850", "GSM 1900", "W-CDMA B2", "W-CDMA B4", "W-CDMA B5", "LTE B2", "LTE B4",
                        "LTE B5", "LTE B7", "LTE B12", "LTE B13", "LTE B14", "LTE B17", "LTE B25",
                        "LTE B30", "LTE B38", "LTE B41 PC3", "LTE B41 PC2", "LTE B53", "LTE B66",
                        "LTE B70", "LTE B71", "FR1 n2", "FR1 n5", "FR1 n7", "FR1 n12", "FR1 n13",
                        "FR1 n14", "FR1 n25", "FR1 n30", "FR1 n38", "FR1 n41 PC3", "FR1 n66",
                        "FR1 n41 PC2", "FR1 n53", "FR1 n78 PC3", "FR1 n78 PC2", "MSS"],
            "DTS" :    ["Wi-Fi 2.4 GHz", "Thread (2.4 GHz)"],
            "NII" :    ["Wi-Fi 5.2 GHz", "Wi-Fi 5.3 GHz", "Wi-Fi 5.5 GHz", "Wi-Fi 5.8 GHz", "Wi-Fi 6E", "U-NII 5",
                        "U-NII 6", "U-NII 7", "U-NII 8", "Bluetooth (NB U-NII 1)", "Bluetooth (NB U-NII 3)"],
            "DSS" :    ["Bluetooth (2.4 GHz)"]
            }
        self.log = log_dir
    
    def section_1_fcc(self):
        try:
            cbe_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["CBE"][tech]] for tech in range(len(self.fcc["CBE"]))]
            pce_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["PCE"][tech]] for tech in range(len(self.fcc["PCE"]))]
            tne_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["TNE"][tech]] for tech in range(len(self.fcc["TNE"]))]
            dts_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["DTS"][tech]] for tech in range(len(self.fcc["DTS"]))]
            nii_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["NII"][tech]] for tech in range(len(self.fcc["NII"]))]
            dss_fcc_data = [self.data[self.data["Technology OG"] == self.fcc["DSS"][tech]] for tech in range(len(self.fcc["DSS"]))]
            
            cbe_fcc_separated_data = [cbe_fcc_data[no_nan][cbe_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(cbe_fcc_data))]
            pce_fcc_separated_data = [pce_fcc_data[no_nan][pce_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(pce_fcc_data))]
            tne_fcc_separated_data = [tne_fcc_data[no_nan][tne_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(tne_fcc_data))]
            dts_fcc_separated_data = [dts_fcc_data[no_nan][dts_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dts_fcc_data))]
            nii_fcc_separated_data = [nii_fcc_data[no_nan][nii_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(nii_fcc_data))]
            dss_fcc_separated_data = [dss_fcc_data[no_nan][dss_fcc_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dss_fcc_data))]
            
            cbe_fcc_head_data = [cbe_fcc_separated_data[head_data][cbe_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(cbe_fcc_separated_data))]
            pce_fcc_head_data = [pce_fcc_separated_data[head_data][pce_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(pce_fcc_separated_data))]
            tne_fcc_head_data = [tne_fcc_separated_data[head_data][tne_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(tne_fcc_separated_data))]
            dts_fcc_head_data = [dts_fcc_separated_data[head_data][dts_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(dts_fcc_separated_data))]
            nii_fcc_head_data = [nii_fcc_separated_data[head_data][nii_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(nii_fcc_separated_data))]
            dss_fcc_head_data = [dss_fcc_separated_data[head_data][dss_fcc_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(dss_fcc_separated_data))]
            
            cbe_fcc_body_data = [cbe_fcc_separated_data[body_data][cbe_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(cbe_fcc_separated_data))]
            pce_fcc_body_data = [pce_fcc_separated_data[body_data][pce_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(pce_fcc_separated_data))]
            tne_fcc_body_data = [tne_fcc_separated_data[body_data][tne_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(tne_fcc_separated_data))]
            dts_fcc_body_data = [dts_fcc_separated_data[body_data][dts_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(dts_fcc_separated_data))]
            nii_fcc_body_data = [nii_fcc_separated_data[body_data][nii_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(nii_fcc_separated_data))]
            dss_fcc_body_data = [dss_fcc_separated_data[body_data][dss_fcc_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(dss_fcc_separated_data))]
            
            cbe_fcc_body_hotspot_data = [cbe_fcc_separated_data[body_hotspot_data][cbe_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(cbe_fcc_separated_data))]
            pce_fcc_body_hotspot_data = [pce_fcc_separated_data[body_hotspot_data][pce_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(pce_fcc_separated_data))]
            tne_fcc_body_hotspot_data = [tne_fcc_separated_data[body_hotspot_data][tne_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(tne_fcc_separated_data))]
            dts_fcc_body_hotspot_data = [dts_fcc_separated_data[body_hotspot_data][dts_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(dts_fcc_separated_data))]
            nii_fcc_body_hotspot_data = [nii_fcc_separated_data[body_hotspot_data][nii_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(nii_fcc_separated_data))]
            dss_fcc_body_hotspot_data = [dss_fcc_separated_data[body_hotspot_data][dss_fcc_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(dss_fcc_separated_data))]
            
            cbe_fcc_hotspot_data = [cbe_fcc_separated_data[hotspot_data][cbe_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(cbe_fcc_separated_data))]
            pce_fcc_hotspot_data = [pce_fcc_separated_data[hotspot_data][pce_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(pce_fcc_separated_data))]
            tne_fcc_hotspot_data = [tne_fcc_separated_data[hotspot_data][tne_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(tne_fcc_separated_data))]
            dts_fcc_hotspot_data = [dts_fcc_separated_data[hotspot_data][dts_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(dts_fcc_separated_data))]
            nii_fcc_hotspot_data = [nii_fcc_separated_data[hotspot_data][nii_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(nii_fcc_separated_data))]
            dss_fcc_hotspot_data = [dss_fcc_separated_data[hotspot_data][dss_fcc_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(dss_fcc_separated_data))]
            
            cbe_fcc_extremity_data = [cbe_fcc_separated_data[extremity_data][cbe_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(cbe_fcc_separated_data))]
            pce_fcc_extremity_data = [pce_fcc_separated_data[extremity_data][pce_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(pce_fcc_separated_data))]
            tne_fcc_extremity_data = [tne_fcc_separated_data[extremity_data][tne_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(tne_fcc_separated_data))]
            dts_fcc_extremity_data = [dts_fcc_separated_data[extremity_data][dts_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(dts_fcc_separated_data))]
            nii_fcc_extremity_data = [nii_fcc_separated_data[extremity_data][nii_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(nii_fcc_separated_data))]
            dss_fcc_extremity_data = [dss_fcc_separated_data[extremity_data][dss_fcc_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(dss_fcc_separated_data))]
            
            cbe_fcc_head_max = [cbe_fcc_head_data[cbe_fcc_head_data_max][cbe_fcc_head_data[cbe_fcc_head_data_max]["1-g Scaled (W/kg)"] == cbe_fcc_head_data[cbe_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for cbe_fcc_head_data_max in range(len(cbe_fcc_head_data))]
            pce_fcc_head_max = [pce_fcc_head_data[pce_fcc_head_data_max][pce_fcc_head_data[pce_fcc_head_data_max]["1-g Scaled (W/kg)"] == pce_fcc_head_data[pce_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for pce_fcc_head_data_max in range(len(pce_fcc_head_data))]
            tne_fcc_head_max = [tne_fcc_head_data[tne_fcc_head_data_max][tne_fcc_head_data[tne_fcc_head_data_max]["1-g Scaled (W/kg)"] == tne_fcc_head_data[tne_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for tne_fcc_head_data_max in range(len(tne_fcc_head_data))]
            dts_fcc_head_max = [dts_fcc_head_data[dts_fcc_head_data_max][dts_fcc_head_data[dts_fcc_head_data_max]["1-g Scaled (W/kg)"] == dts_fcc_head_data[dts_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for dts_fcc_head_data_max in range(len(dts_fcc_head_data))]
            nii_fcc_head_max = [nii_fcc_head_data[nii_fcc_head_data_max][nii_fcc_head_data[nii_fcc_head_data_max]["1-g Scaled (W/kg)"] == nii_fcc_head_data[nii_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for nii_fcc_head_data_max in range(len(nii_fcc_head_data))]
            dss_fcc_head_max = [dss_fcc_head_data[dss_fcc_head_data_max][dss_fcc_head_data[dss_fcc_head_data_max]["1-g Scaled (W/kg)"] == dss_fcc_head_data[dss_fcc_head_data_max]["1-g Scaled (W/kg)"].max()] for dss_fcc_head_data_max in range(len(dss_fcc_head_data))]
            
            cbe_fcc_body_max = [cbe_fcc_body_data[cbe_fcc_body_data_max][cbe_fcc_body_data[cbe_fcc_body_data_max]["1-g Scaled (W/kg)"] == cbe_fcc_body_data[cbe_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for cbe_fcc_body_data_max in range(len(cbe_fcc_body_data))]
            pce_fcc_body_max = [pce_fcc_body_data[pce_fcc_body_data_max][pce_fcc_body_data[pce_fcc_body_data_max]["1-g Scaled (W/kg)"] == pce_fcc_body_data[pce_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for pce_fcc_body_data_max in range(len(pce_fcc_body_data))]
            tne_fcc_body_max = [tne_fcc_body_data[tne_fcc_body_data_max][tne_fcc_body_data[tne_fcc_body_data_max]["1-g Scaled (W/kg)"] == tne_fcc_body_data[tne_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for tne_fcc_body_data_max in range(len(tne_fcc_body_data))]
            dts_fcc_body_max = [dts_fcc_body_data[dts_fcc_body_data_max][dts_fcc_body_data[dts_fcc_body_data_max]["1-g Scaled (W/kg)"] == dts_fcc_body_data[dts_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for dts_fcc_body_data_max in range(len(dts_fcc_body_data))]
            nii_fcc_body_max = [nii_fcc_body_data[nii_fcc_body_data_max][nii_fcc_body_data[nii_fcc_body_data_max]["1-g Scaled (W/kg)"] == nii_fcc_body_data[nii_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for nii_fcc_body_data_max in range(len(nii_fcc_body_data))]
            dss_fcc_body_max = [dss_fcc_body_data[dss_fcc_body_data_max][dss_fcc_body_data[dss_fcc_body_data_max]["1-g Scaled (W/kg)"] == dss_fcc_body_data[dss_fcc_body_data_max]["1-g Scaled (W/kg)"].max()] for dss_fcc_body_data_max in range(len(dss_fcc_body_data))]
            
            cbe_fcc_body_hotspot_max = [cbe_fcc_body_hotspot_data[cbe_fcc_body_hotspot_data_max][cbe_fcc_body_hotspot_data[cbe_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == cbe_fcc_body_hotspot_data[cbe_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for cbe_fcc_body_hotspot_data_max in range(len(cbe_fcc_body_hotspot_data))]
            pce_fcc_body_hotspot_max = [pce_fcc_body_hotspot_data[pce_fcc_body_hotspot_data_max][pce_fcc_body_hotspot_data[pce_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == pce_fcc_body_hotspot_data[pce_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for pce_fcc_body_hotspot_data_max in range(len(pce_fcc_body_hotspot_data))]
            tne_fcc_body_hotspot_max = [tne_fcc_body_hotspot_data[tne_fcc_body_hotspot_data_max][tne_fcc_body_hotspot_data[tne_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == tne_fcc_body_hotspot_data[tne_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for tne_fcc_body_hotspot_data_max in range(len(tne_fcc_body_hotspot_data))]
            dts_fcc_body_hotspot_max = [dts_fcc_body_hotspot_data[dts_fcc_body_hotspot_data_max][dts_fcc_body_hotspot_data[dts_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == dts_fcc_body_hotspot_data[dts_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dts_fcc_body_hotspot_data_max in range(len(dts_fcc_body_hotspot_data))]
            nii_fcc_body_hotspot_max = [nii_fcc_body_hotspot_data[nii_fcc_body_hotspot_data_max][nii_fcc_body_hotspot_data[nii_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == nii_fcc_body_hotspot_data[nii_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for nii_fcc_body_hotspot_data_max in range(len(nii_fcc_body_hotspot_data))]
            dss_fcc_body_hotspot_max = [dss_fcc_body_hotspot_data[dss_fcc_body_hotspot_data_max][dss_fcc_body_hotspot_data[dss_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"] == dss_fcc_body_hotspot_data[dss_fcc_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dss_fcc_body_hotspot_data_max in range(len(dss_fcc_body_hotspot_data))]
            
            cbe_fcc_hotspot_max = [cbe_fcc_hotspot_data[cbe_fcc_hotspot_data_max][cbe_fcc_hotspot_data[cbe_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == cbe_fcc_hotspot_data[cbe_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for cbe_fcc_hotspot_data_max in range(len(cbe_fcc_hotspot_data))]
            pce_fcc_hotspot_max = [pce_fcc_hotspot_data[pce_fcc_hotspot_data_max][pce_fcc_hotspot_data[pce_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == pce_fcc_hotspot_data[pce_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for pce_fcc_hotspot_data_max in range(len(pce_fcc_hotspot_data))]
            tne_fcc_hotspot_max = [tne_fcc_hotspot_data[tne_fcc_hotspot_data_max][tne_fcc_hotspot_data[tne_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == tne_fcc_hotspot_data[tne_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for tne_fcc_hotspot_data_max in range(len(tne_fcc_hotspot_data))]
            dts_fcc_hotspot_max = [dts_fcc_hotspot_data[dts_fcc_hotspot_data_max][dts_fcc_hotspot_data[dts_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == dts_fcc_hotspot_data[dts_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dts_fcc_hotspot_data_max in range(len(dts_fcc_hotspot_data))]
            nii_fcc_hotspot_max = [nii_fcc_hotspot_data[nii_fcc_hotspot_data_max][nii_fcc_hotspot_data[nii_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == nii_fcc_hotspot_data[nii_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for nii_fcc_hotspot_data_max in range(len(nii_fcc_hotspot_data))]
            dss_fcc_hotspot_max = [dss_fcc_hotspot_data[dss_fcc_hotspot_data_max][dss_fcc_hotspot_data[dss_fcc_hotspot_data_max]["1-g Scaled (W/kg)"] == dss_fcc_hotspot_data[dss_fcc_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dss_fcc_hotspot_data_max in range(len(dss_fcc_hotspot_data))]
            
            cbe_fcc_extremity_max = [cbe_fcc_extremity_data[cbe_fcc_extremity_data_max][cbe_fcc_extremity_data[cbe_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == cbe_fcc_extremity_data[cbe_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for cbe_fcc_extremity_data_max in range(len(cbe_fcc_extremity_data))]
            pce_fcc_extremity_max = [pce_fcc_extremity_data[pce_fcc_extremity_data_max][pce_fcc_extremity_data[pce_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == pce_fcc_extremity_data[pce_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for pce_fcc_extremity_data_max in range(len(pce_fcc_extremity_data))]
            tne_fcc_extremity_max = [tne_fcc_extremity_data[tne_fcc_extremity_data_max][tne_fcc_extremity_data[tne_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == tne_fcc_extremity_data[tne_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for tne_fcc_extremity_data_max in range(len(tne_fcc_extremity_data))]
            dts_fcc_extremity_max = [dts_fcc_extremity_data[dts_fcc_extremity_data_max][dts_fcc_extremity_data[dts_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == dts_fcc_extremity_data[dts_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for dts_fcc_extremity_data_max in range(len(dts_fcc_extremity_data))]
            nii_fcc_extremity_max = [nii_fcc_extremity_data[nii_fcc_extremity_data_max][nii_fcc_extremity_data[nii_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == nii_fcc_extremity_data[nii_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for nii_fcc_extremity_data_max in range(len(nii_fcc_extremity_data))]
            dss_fcc_extremity_max = [dss_fcc_extremity_data[dss_fcc_extremity_data_max][dss_fcc_extremity_data[dss_fcc_extremity_data_max]["10-g Scaled (W/kg)"] == dss_fcc_extremity_data[dss_fcc_extremity_data_max]["10-g Scaled (W/kg)"].max()] for dss_fcc_extremity_data_max in range(len(dss_fcc_extremity_data))]
            
            cbe_fcc_sec_1 = [pd.concat(cbe_fcc_head_max), pd.concat(cbe_fcc_body_max), pd.concat(cbe_fcc_body_hotspot_max), pd.concat(cbe_fcc_hotspot_max), pd.concat(cbe_fcc_extremity_max)]
            pce_fcc_sec_1 = [pd.concat(pce_fcc_head_max), pd.concat(pce_fcc_body_max), pd.concat(pce_fcc_body_hotspot_max), pd.concat(pce_fcc_hotspot_max), pd.concat(pce_fcc_extremity_max)]
            tne_fcc_sec_1 = [pd.concat(tne_fcc_head_max), pd.concat(tne_fcc_body_max), pd.concat(tne_fcc_body_hotspot_max), pd.concat(tne_fcc_hotspot_max), pd.concat(tne_fcc_extremity_max)]
            dts_fcc_sec_1 = [pd.concat(dts_fcc_head_max), pd.concat(dts_fcc_body_max), pd.concat(dts_fcc_body_hotspot_max), pd.concat(dts_fcc_hotspot_max), pd.concat(dts_fcc_extremity_max)]
            nii_fcc_sec_1 = [pd.concat(nii_fcc_head_max), pd.concat(nii_fcc_body_max), pd.concat(nii_fcc_body_hotspot_max), pd.concat(nii_fcc_hotspot_max), pd.concat(nii_fcc_extremity_max)]
            dss_fcc_sec_1 = [pd.concat(dss_fcc_head_max), pd.concat(dss_fcc_body_max), pd.concat(dss_fcc_body_hotspot_max), pd.concat(dss_fcc_hotspot_max), pd.concat(dss_fcc_extremity_max)]
            
            if cbe_fcc_sec_1[0].size == 0:
                cbe_head = "N/A"
            else:
                cbe_head_max = cbe_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                cbe_head = float(cbe_head_max)
            if cbe_fcc_sec_1[1].size == 0:
                cbe_body = "N/A"
            else:
                cbe_body_max = cbe_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                cbe_body = float(cbe_body_max)
            if cbe_fcc_sec_1[2].size == 0:
                cbe_body_hoty = "N/A"
            else:
                cbe_body_hoty_max = cbe_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                cbe_body_hoty = float(cbe_body_hoty_max)
            if cbe_fcc_sec_1[3].size == 0:
                cbe_hoty = "N/A"
            else:
                cbe_hoty_max = cbe_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                cbe_hoty = float(cbe_hoty_max)
            if cbe_fcc_sec_1[4].size == 0:
                cbe_extr = "N/A"
            else:
                cbe_extr_max = cbe_fcc_sec_1[4]["1-g Scaled (W/kg)"].max()
                cbe_extr = float(cbe_extr_max)
            
            if pce_fcc_sec_1[0].size == 0:
                pce_head = "N/A"
            else:
                pce_head_max = pce_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                pce_head = float(pce_head_max)
            if pce_fcc_sec_1[1].size == 0:
                pce_body = "N/A"
            else:
                pce_body_max = pce_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                pce_body = float(pce_body_max)
            if pce_fcc_sec_1[2].size == 0:
                pce_body_hoty = "N/A"
            else:
                pce_body_hoty_max = pce_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                pce_body_hoty = float(pce_body_hoty_max)
            if pce_fcc_sec_1[3].size == 0:
                pce_hoty = "N/A"
            else:
                pce_hoty_max = pce_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                pce_hoty = float(pce_hoty_max)
            if pce_fcc_sec_1[4].size == 0:
                pce_extr = "N/A"
            else:
                pce_extr_max = pce_fcc_sec_1[4]["10-g Scaled (W/kg)"].max()
                pce_extr = float(pce_extr_max)
            
            if tne_fcc_sec_1[0].size == 0:
                tne_head = "N/A"
            else:
                tne_head_max = tne_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                tne_head = float(tne_head_max)
            if tne_fcc_sec_1[1].size == 0:
                tne_body = "N/A"
            else:
                tne_body_max = tne_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                tne_body = float(tne_body_max)
            if tne_fcc_sec_1[2].size == 0:
                tne_body_hoty = "N/A"
            else:
                tne_body_hoty_max = tne_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                tne_body_hoty = float(tne_body_hoty_max)
            if tne_fcc_sec_1[3].size == 0:
                tne_hoty = "N/A"
            else:
                tne_hoty_max = tne_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                tne_hoty = float(tne_hoty_max)
            if tne_fcc_sec_1[4].size == 0:
                tne_extr = "N/A"
            else:
                tne_extr_max = tne_fcc_sec_1[4]["10-g Scaled (W/kg)"].max()
                tne_extr = float(tne_extr_max)
            
            if dts_fcc_sec_1[0].size == 0:
                dts_head = "N/A"
            else:
                dts_head_max = dts_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                dts_head = float(dts_head_max)
            if dts_fcc_sec_1[1].size == 0:
                dts_body = "N/A"
            else:
                dts_body_max = dts_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                dts_body = float(dts_body_max)
            if dts_fcc_sec_1[2].size == 0:
                dts_body_hoty = "N/A"
            else:
                dts_body_hoty_max = dts_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                dts_body_hoty = float(dts_body_hoty_max)
            if dts_fcc_sec_1[3].size == 0:
                dts_hoty = "N/A"
            else:
                dts_hoty_max = dts_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                dts_hoty = float(dts_hoty_max)
            if dts_fcc_sec_1[4].size == 0:
                dts_extr = "N/A"
            else:
                dts_extr_max = dts_fcc_sec_1[4]["10-g Scaled (W/kg)"].max()
                dts_extr = float(dts_extr_max)
            
            if nii_fcc_sec_1[0].size == 0:
                nii_head = "N/A"
            else:
                nii_head_max = nii_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                nii_head = float(nii_head_max)
            if nii_fcc_sec_1[1].size == 0:
                nii_body = "N/A"
            else:
                nii_body_max = nii_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                nii_body = float(nii_body_max)
            if nii_fcc_sec_1[2].size == 0:
                nii_body_hoty = "N/A"
            else:
                nii_body_hoty_max = nii_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                nii_body_hoty = float(nii_body_hoty_max)
            if nii_fcc_sec_1[3].size == 0:
                nii_hoty = "N/A"
            else:
                nii_hoty_max = nii_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                nii_hoty = float(nii_hoty_max)
            if nii_fcc_sec_1[4].size == 0:
                nii_extr = "N/A"
            else:
                nii_extr_max = nii_fcc_sec_1[4]["10-g Scaled (W/kg)"].max()
                nii_extr = float(nii_extr_max)
            
            if dss_fcc_sec_1[0].size == 0:
                dss_head = "N/A"
            else:
                dss_head_max = dss_fcc_sec_1[0]["1-g Scaled (W/kg)"].max()
                dss_head = float(dss_head_max)
            if dss_fcc_sec_1[1].size == 0:
                dss_body = "N/A"
            else:
                dss_body_max = dss_fcc_sec_1[1]["1-g Scaled (W/kg)"].max()
                dss_body = float(dss_body_max)
            if dss_fcc_sec_1[2].size == 0:
                dss_body_hoty = "N/A"
            else:
                dss_body_hoty_max = dss_fcc_sec_1[2]["1-g Scaled (W/kg)"].max()
                dss_body_hoty = float(dss_body_hoty_max)
            if dss_fcc_sec_1[3].size == 0:
                dss_hoty = "N/A"
            else:
                dss_hoty_max = dss_fcc_sec_1[3]["1-g Scaled (W/kg)"].max()
                dss_hoty = float(dss_hoty_max)
            if dss_fcc_sec_1[4].size == 0:
                dss_extr = "N/A"
            else:
                dss_extr_max = dss_fcc_sec_1[4]["10-g Scaled (W/kg)"].max()
                dss_extr = float(dss_extr_max)
            
            section_1_summary = {
                "RF Exposure Condition": [
                    "Head",
                    "Body-worn",
                    "Body & Hotspot",
                    "Hotspot",
                    "Extremity"
                    ],
                "TNE": [
                    tne_head,
                    tne_body,
                    tne_body_hoty,
                    tne_hoty,
                    tne_extr
                    ],
                "PCE": [
                    pce_head,
                    pce_body,
                    pce_body_hoty,
                    pce_hoty,
                    pce_extr
                    ],
                "CBE": [
                    cbe_head,
                    cbe_body,
                    cbe_body_hoty,
                    cbe_hoty,
                    cbe_extr
                    ],
                "DTS": [
                    dts_head,
                    dts_body,
                    dts_body_hoty,
                    dts_hoty,
                    dts_extr
                    ],
                "NII": [
                    nii_head,
                    nii_body,
                    nii_body_hoty,
                    nii_hoty,
                    nii_extr
                    ],
                "DSS": [
                    dss_head,
                    dss_body,
                    dss_body_hoty,
                    dss_hoty,
                    dss_extr
                    ],
                }
            
            section_1_summary_df = pd.DataFrame(section_1_summary)
            
            cbe_fcc_max_final = pd.concat([pd.concat(cbe_fcc_head_max), pd.concat(cbe_fcc_body_max), pd.concat(cbe_fcc_body_hotspot_max), pd.concat(cbe_fcc_hotspot_max), pd.concat(cbe_fcc_extremity_max)]).reset_index()
            pce_fcc_max_final = pd.concat([pd.concat(pce_fcc_head_max), pd.concat(pce_fcc_body_max), pd.concat(pce_fcc_body_hotspot_max), pd.concat(pce_fcc_hotspot_max), pd.concat(pce_fcc_extremity_max)]).reset_index()
            tne_fcc_max_final = pd.concat([pd.concat(tne_fcc_head_max), pd.concat(tne_fcc_body_max), pd.concat(tne_fcc_body_hotspot_max), pd.concat(tne_fcc_hotspot_max), pd.concat(tne_fcc_extremity_max)]).reset_index()
            dts_fcc_max_final = pd.concat([pd.concat(dts_fcc_head_max), pd.concat(dts_fcc_body_max), pd.concat(dts_fcc_body_hotspot_max), pd.concat(dts_fcc_hotspot_max), pd.concat(dts_fcc_extremity_max)]).reset_index()
            nii_fcc_max_final = pd.concat([pd.concat(nii_fcc_head_max), pd.concat(nii_fcc_body_max), pd.concat(nii_fcc_body_hotspot_max), pd.concat(nii_fcc_hotspot_max), pd.concat(nii_fcc_extremity_max)]).reset_index()
            dss_fcc_max_final = pd.concat([pd.concat(dss_fcc_head_max), pd.concat(dss_fcc_body_max), pd.concat(dss_fcc_body_hotspot_max), pd.concat(dss_fcc_hotspot_max), pd.concat(dss_fcc_extremity_max)]).reset_index()
            
            cbe_fcc_max_final_filtered = cbe_fcc_max_final.loc[cbe_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            pce_fcc_max_final_filtered = pce_fcc_max_final.loc[pce_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            tne_fcc_max_final_filtered = tne_fcc_max_final.loc[tne_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            dts_fcc_max_final_filtered = dts_fcc_max_final.loc[dts_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            nii_fcc_max_final_filtered = nii_fcc_max_final.loc[nii_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            dss_fcc_max_final_filtered = dss_fcc_max_final.loc[dss_fcc_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            cbe_final = cbe_fcc_max_final_filtered
            pce_final = pce_fcc_max_final_filtered
            tne_final = tne_fcc_max_final_filtered
            dts_final = dts_fcc_max_final_filtered
            nii_final = nii_fcc_max_final_filtered
            dss_final = dss_fcc_max_final_filtered
            
            final_list = [section_1_summary_df, tne_final, pce_final, cbe_final, dts_final, nii_final, dss_final]
            
            equip_list = ["Section 1 Summary", "TNE", "PCE", "CBE", "DTS", "NII", "DSS"]
            
            try:
                abs_filepath = Path(f"{self.s1_fcc}").resolve(strict = True)
            
            except FileNotFoundError:
                with pd.ExcelWriter(f"{self.s1_fcc}") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(final_list)):  
                        final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
            
            else:
                with pd.ExcelWriter(f"{self.s1_fcc}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(final_list)):  
                        final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
        
        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Section 1 FCC Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)
    
    def section_1_ised(self):
        
        try:
            pce_ised_data = [self.data[self.data["Technology OG"] == self.ised["PCE"][tech]] for tech in range(len(self.ised["PCE"]))]
            dts_ised_data = [self.data[self.data["Technology OG"] == self.ised["DTS"][tech]] for tech in range(len(self.ised["DTS"]))]
            nii_ised_data = [self.data[self.data["Technology OG"] == self.ised["NII"][tech]] for tech in range(len(self.ised["NII"]))]
            dss_ised_data = [self.data[self.data["Technology OG"] == self.ised["DSS"][tech]] for tech in range(len(self.ised["DSS"]))]
            
            pce_ised_separated_data = [pce_ised_data[no_nan][pce_ised_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(pce_ised_data))]
            dts_ised_separated_data = [dts_ised_data[no_nan][dts_ised_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dts_ised_data))]
            nii_ised_separated_data = [nii_ised_data[no_nan][nii_ised_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(nii_ised_data))]
            dss_ised_separated_data = [dss_ised_data[no_nan][dss_ised_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dss_ised_data))]
            
            pce_ised_head_data = [pce_ised_separated_data[head_data][pce_ised_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(pce_ised_separated_data))]
            dts_ised_head_data = [dts_ised_separated_data[head_data][dts_ised_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(dts_ised_separated_data))]
            nii_ised_head_data = [nii_ised_separated_data[head_data][nii_ised_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(nii_ised_separated_data))]
            dss_ised_head_data = [dss_ised_separated_data[head_data][dss_ised_separated_data[head_data]["RF Exposure Condition"] == "Head"] for head_data in range(len(dss_ised_separated_data))]
            
            pce_ised_body_data = [pce_ised_separated_data[body_data][pce_ised_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(pce_ised_separated_data))]
            dts_ised_body_data = [dts_ised_separated_data[body_data][dts_ised_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(dts_ised_separated_data))]
            nii_ised_body_data = [nii_ised_separated_data[body_data][nii_ised_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(nii_ised_separated_data))]
            dss_ised_body_data = [dss_ised_separated_data[body_data][dss_ised_separated_data[body_data]["RF Exposure Condition"] == "Body-worn"] for body_data in range(len(dss_ised_separated_data))]
            
            pce_ised_body_hotspot_data = [pce_ised_separated_data[body_hotspot_data][pce_ised_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(pce_ised_separated_data))]
            dts_ised_body_hotspot_data = [dts_ised_separated_data[body_hotspot_data][dts_ised_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(dts_ised_separated_data))]
            nii_ised_body_hotspot_data = [nii_ised_separated_data[body_hotspot_data][nii_ised_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(nii_ised_separated_data))]
            dss_ised_body_hotspot_data = [dss_ised_separated_data[body_hotspot_data][dss_ised_separated_data[body_hotspot_data]["RF Exposure Condition"] == "Body & Hotspot"] for body_hotspot_data in range(len(dss_ised_separated_data))]
            
            pce_ised_hotspot_data = [pce_ised_separated_data[hotspot_data][pce_ised_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(pce_ised_separated_data))]
            dts_ised_hotspot_data = [dts_ised_separated_data[hotspot_data][dts_ised_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(dts_ised_separated_data))]
            nii_ised_hotspot_data = [nii_ised_separated_data[hotspot_data][nii_ised_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(nii_ised_separated_data))]
            dss_ised_hotspot_data = [dss_ised_separated_data[hotspot_data][dss_ised_separated_data[hotspot_data]["RF Exposure Condition"] == "Hotspot"] for hotspot_data in range(len(dss_ised_separated_data))]
            
            pce_ised_extremity_data = [pce_ised_separated_data[extremity_data][pce_ised_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(pce_ised_separated_data))]
            dts_ised_extremity_data = [dts_ised_separated_data[extremity_data][dts_ised_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(dts_ised_separated_data))]
            nii_ised_extremity_data = [nii_ised_separated_data[extremity_data][nii_ised_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(nii_ised_separated_data))]
            dss_ised_extremity_data = [dss_ised_separated_data[extremity_data][dss_ised_separated_data[extremity_data]["RF Exposure Condition"] == "Extremity"] for extremity_data in range(len(dss_ised_separated_data))]
            
            pce_ised_head_max = [pce_ised_head_data[pce_ised_head_data_max][pce_ised_head_data[pce_ised_head_data_max]["1-g Scaled (W/kg)"] == pce_ised_head_data[pce_ised_head_data_max]["1-g Scaled (W/kg)"].max()] for pce_ised_head_data_max in range(len(pce_ised_head_data))]
            dts_ised_head_max = [dts_ised_head_data[dts_ised_head_data_max][dts_ised_head_data[dts_ised_head_data_max]["1-g Scaled (W/kg)"] == dts_ised_head_data[dts_ised_head_data_max]["1-g Scaled (W/kg)"].max()] for dts_ised_head_data_max in range(len(dts_ised_head_data))]
            nii_ised_head_max = [nii_ised_head_data[nii_ised_head_data_max][nii_ised_head_data[nii_ised_head_data_max]["1-g Scaled (W/kg)"] == nii_ised_head_data[nii_ised_head_data_max]["1-g Scaled (W/kg)"].max()] for nii_ised_head_data_max in range(len(nii_ised_head_data))]
            dss_ised_head_max = [dss_ised_head_data[dss_ised_head_data_max][dss_ised_head_data[dss_ised_head_data_max]["1-g Scaled (W/kg)"] == dss_ised_head_data[dss_ised_head_data_max]["1-g Scaled (W/kg)"].max()] for dss_ised_head_data_max in range(len(dss_ised_head_data))]
            
            pce_ised_body_max = [pce_ised_body_data[pce_ised_body_data_max][pce_ised_body_data[pce_ised_body_data_max]["1-g Scaled (W/kg)"] == pce_ised_body_data[pce_ised_body_data_max]["1-g Scaled (W/kg)"].max()] for pce_ised_body_data_max in range(len(pce_ised_body_data))]
            dts_ised_body_max = [dts_ised_body_data[dts_ised_body_data_max][dts_ised_body_data[dts_ised_body_data_max]["1-g Scaled (W/kg)"] == dts_ised_body_data[dts_ised_body_data_max]["1-g Scaled (W/kg)"].max()] for dts_ised_body_data_max in range(len(dts_ised_body_data))]
            nii_ised_body_max = [nii_ised_body_data[nii_ised_body_data_max][nii_ised_body_data[nii_ised_body_data_max]["1-g Scaled (W/kg)"] == nii_ised_body_data[nii_ised_body_data_max]["1-g Scaled (W/kg)"].max()] for nii_ised_body_data_max in range(len(nii_ised_body_data))]
            dss_ised_body_max = [dss_ised_body_data[dss_ised_body_data_max][dss_ised_body_data[dss_ised_body_data_max]["1-g Scaled (W/kg)"] == dss_ised_body_data[dss_ised_body_data_max]["1-g Scaled (W/kg)"].max()] for dss_ised_body_data_max in range(len(dss_ised_body_data))]
            
            pce_ised_body_hotspot_max = [pce_ised_body_hotspot_data[pce_ised_body_hotspot_data_max][pce_ised_body_hotspot_data[pce_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"] == pce_ised_body_hotspot_data[pce_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for pce_ised_body_hotspot_data_max in range(len(pce_ised_body_hotspot_data))]
            dts_ised_body_hotspot_max = [dts_ised_body_hotspot_data[dts_ised_body_hotspot_data_max][dts_ised_body_hotspot_data[dts_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"] == dts_ised_body_hotspot_data[dts_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dts_ised_body_hotspot_data_max in range(len(dts_ised_body_hotspot_data))]
            nii_ised_body_hotspot_max = [nii_ised_body_hotspot_data[nii_ised_body_hotspot_data_max][nii_ised_body_hotspot_data[nii_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"] == nii_ised_body_hotspot_data[nii_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for nii_ised_body_hotspot_data_max in range(len(nii_ised_body_hotspot_data))]
            dss_ised_body_hotspot_max = [dss_ised_body_hotspot_data[dss_ised_body_hotspot_data_max][dss_ised_body_hotspot_data[dss_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"] == dss_ised_body_hotspot_data[dss_ised_body_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dss_ised_body_hotspot_data_max in range(len(dss_ised_body_hotspot_data))]
            
            pce_ised_hotspot_max = [pce_ised_hotspot_data[pce_ised_hotspot_data_max][pce_ised_hotspot_data[pce_ised_hotspot_data_max]["1-g Scaled (W/kg)"] == pce_ised_hotspot_data[pce_ised_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for pce_ised_hotspot_data_max in range(len(pce_ised_hotspot_data))]
            dts_ised_hotspot_max = [dts_ised_hotspot_data[dts_ised_hotspot_data_max][dts_ised_hotspot_data[dts_ised_hotspot_data_max]["1-g Scaled (W/kg)"] == dts_ised_hotspot_data[dts_ised_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dts_ised_hotspot_data_max in range(len(dts_ised_hotspot_data))]
            nii_ised_hotspot_max = [nii_ised_hotspot_data[nii_ised_hotspot_data_max][nii_ised_hotspot_data[nii_ised_hotspot_data_max]["1-g Scaled (W/kg)"] == nii_ised_hotspot_data[nii_ised_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for nii_ised_hotspot_data_max in range(len(nii_ised_hotspot_data))]
            dss_ised_hotspot_max = [dss_ised_hotspot_data[dss_ised_hotspot_data_max][dss_ised_hotspot_data[dss_ised_hotspot_data_max]["1-g Scaled (W/kg)"] == dss_ised_hotspot_data[dss_ised_hotspot_data_max]["1-g Scaled (W/kg)"].max()] for dss_ised_hotspot_data_max in range(len(dss_ised_hotspot_data))]
            
            pce_ised_extremity_max = [pce_ised_extremity_data[pce_ised_extremity_data_max][pce_ised_extremity_data[pce_ised_extremity_data_max]["10-g Scaled (W/kg)"] == pce_ised_extremity_data[pce_ised_extremity_data_max]["10-g Scaled (W/kg)"].max()] for pce_ised_extremity_data_max in range(len(pce_ised_extremity_data))]
            dts_ised_extremity_max = [dts_ised_extremity_data[dts_ised_extremity_data_max][dts_ised_extremity_data[dts_ised_extremity_data_max]["10-g Scaled (W/kg)"] == dts_ised_extremity_data[dts_ised_extremity_data_max]["10-g Scaled (W/kg)"].max()] for dts_ised_extremity_data_max in range(len(dts_ised_extremity_data))]
            nii_ised_extremity_max = [nii_ised_extremity_data[nii_ised_extremity_data_max][nii_ised_extremity_data[nii_ised_extremity_data_max]["10-g Scaled (W/kg)"] == nii_ised_extremity_data[nii_ised_extremity_data_max]["10-g Scaled (W/kg)"].max()] for nii_ised_extremity_data_max in range(len(nii_ised_extremity_data))]
            dss_ised_extremity_max = [dss_ised_extremity_data[dss_ised_extremity_data_max][dss_ised_extremity_data[dss_ised_extremity_data_max]["10-g Scaled (W/kg)"] == dss_ised_extremity_data[dss_ised_extremity_data_max]["10-g Scaled (W/kg)"].max()] for dss_ised_extremity_data_max in range(len(dss_ised_extremity_data))]
            
            pce_ised_sec_1 = [pd.concat(pce_ised_head_max), pd.concat(pce_ised_body_max), pd.concat(pce_ised_body_hotspot_max), pd.concat(pce_ised_hotspot_max), pd.concat(pce_ised_extremity_max)]
            dts_ised_sec_1 = [pd.concat(dts_ised_head_max), pd.concat(dts_ised_body_max), pd.concat(dts_ised_body_hotspot_max), pd.concat(dts_ised_hotspot_max), pd.concat(dts_ised_extremity_max)]
            nii_ised_sec_1 = [pd.concat(nii_ised_head_max), pd.concat(nii_ised_body_max), pd.concat(nii_ised_body_hotspot_max), pd.concat(nii_ised_hotspot_max), pd.concat(nii_ised_extremity_max)]
            dss_ised_sec_1 = [pd.concat(dss_ised_head_max), pd.concat(dss_ised_body_max), pd.concat(dss_ised_body_hotspot_max), pd.concat(dss_ised_hotspot_max), pd.concat(dss_ised_extremity_max)]
            
            if pce_ised_sec_1[0].size == 0:
                pce_head = "N/A"
            else:
                pce_head_max = pce_ised_sec_1[0]["1-g Scaled (W/kg)"].max()
                pce_head = float(pce_head_max)
            if pce_ised_sec_1[1].size == 0:
                pce_body = "N/A"
            else:
                pce_body_max = pce_ised_sec_1[1]["1-g Scaled (W/kg)"].max()
                pce_body = float(pce_body_max)
            if pce_ised_sec_1[2].size == 0:
                pce_body_hoty = "N/A"
            else:
                pce_body_hoty_max = pce_ised_sec_1[2]["1-g Scaled (W/kg)"].max()
                pce_body_hoty = float(pce_body_hoty_max)
            if pce_ised_sec_1[3].size == 0:
                pce_hoty = "N/A"
            else:
                pce_hoty_max = pce_ised_sec_1[3]["1-g Scaled (W/kg)"].max()
                pce_hoty = float(pce_hoty_max)
            if pce_ised_sec_1[4].size == 0:
                pce_extr = "N/A"
            else:
                pce_extr_max = pce_ised_sec_1[4]["10-g Scaled (W/kg)"].max()
                pce_extr = float(pce_extr_max)
            
            if dts_ised_sec_1[0].size == 0:
                dts_head = "N/A"
            else:
                dts_head_max = dts_ised_sec_1[0]["1-g Scaled (W/kg)"].max()
                dts_head = float(dts_head_max)
            if dts_ised_sec_1[1].size == 0:
                dts_body = "N/A"
            else:
                dts_body_max = dts_ised_sec_1[1]["1-g Scaled (W/kg)"].max()
                dts_body = float(dts_body_max)
            if dts_ised_sec_1[2].size == 0:
                dts_body_hoty = "N/A"
            else:
                dts_body_hoty_max = dts_ised_sec_1[2]["1-g Scaled (W/kg)"].max()
                dts_body_hoty = float(dts_body_hoty_max)
            if dts_ised_sec_1[3].size == 0:
                dts_hoty = "N/A"
            else:
                dts_hoty_max = dts_ised_sec_1[3]["1-g Scaled (W/kg)"].max()
                dts_hoty = float(dts_hoty_max)
            if dts_ised_sec_1[4].size == 0:
                dts_extr = "N/A"
            else:
                dts_extr_max = dts_ised_sec_1[4]["10-g Scaled (W/kg)"].max()
                dts_extr = float(dts_extr_max)
            
            if nii_ised_sec_1[0].size == 0:
                nii_head = "N/A"
            else:
                nii_head_max = nii_ised_sec_1[0]["1-g Scaled (W/kg)"].max()
                nii_head = float(nii_head_max)
            if nii_ised_sec_1[1].size == 0:
                nii_body = "N/A"
            else:
                nii_body_max = nii_ised_sec_1[1]["1-g Scaled (W/kg)"].max()
                nii_body = float(nii_body_max)
            if nii_ised_sec_1[2].size == 0:
                nii_body_hoty = "N/A"
            else:
                nii_body_hoty_max = nii_ised_sec_1[2]["1-g Scaled (W/kg)"].max()
                nii_body_hoty = float(nii_body_hoty_max)
            if nii_ised_sec_1[3].size == 0:
                nii_hoty = "N/A"
            else:
                nii_hoty_max = nii_ised_sec_1[3]["1-g Scaled (W/kg)"].max()
                nii_hoty = float(nii_hoty_max)
            if nii_ised_sec_1[4].size == 0:
                nii_extr = "N/A"
            else:
                nii_extr_max = nii_ised_sec_1[4]["10-g Scaled (W/kg)"].max()
                nii_extr = float(nii_extr_max)
            
            if dss_ised_sec_1[0].size == 0:
                dss_head = "N/A"
            else:
                dss_head_max = dss_ised_sec_1[0]["1-g Scaled (W/kg)"].max()
                dss_head = float(dss_head_max)
            if dss_ised_sec_1[1].size == 0:
                dss_body = "N/A"
            else:
                dss_body_max = dss_ised_sec_1[1]["1-g Scaled (W/kg)"].max()
                dss_body = float(dss_body_max)
            if dss_ised_sec_1[2].size == 0:
                dss_body_hoty = "N/A"
            else:
                dss_body_hoty_max = dss_ised_sec_1[2]["1-g Scaled (W/kg)"].max()
                dss_body_hoty = float(dss_body_hoty_max)
            if dss_ised_sec_1[3].size == 0:
                dss_hoty = "N/A"
            else:
                dss_hoty_max = dss_ised_sec_1[3]["1-g Scaled (W/kg)"].max()
                dss_hoty = float(dss_hoty_max)
            if dss_ised_sec_1[4].size == 0:
                dss_extr = "N/A"
            else:
                dss_extr_max = dss_ised_sec_1[4]["10-g Scaled (W/kg)"].max()
                dss_extr = float(dss_extr_max)
            
            section_1_summary = {
                "RF Exposure Condition":
                    ["Head",
                    "Body-worn",
                    "Body & Hotspot",
                    "Hotspot",
                    "Extremity"],
                "PCE":
                    [pce_head,
                    pce_body,
                    pce_body_hoty,
                    pce_hoty,
                    pce_extr],
                "DTS":
                    [dts_head,
                    dts_body,
                    dts_body_hoty,
                    dts_hoty,
                    dts_extr],
                "NII":
                    [nii_head,
                    nii_body,
                    nii_body_hoty,
                    nii_hoty,
                    nii_extr],
                "DSS":
                    [dss_head,
                    dss_body,
                    dss_body_hoty,
                    dss_hoty,
                    dss_extr],
                    }
            
            section_1_summary_df = pd.DataFrame(section_1_summary)
            
            pce_ised_max_final = pd.concat([pd.concat(pce_ised_head_max), pd.concat(pce_ised_body_max), pd.concat(pce_ised_body_hotspot_max), pd.concat(pce_ised_hotspot_max), pd.concat(pce_ised_extremity_max)]).reset_index()
            dts_ised_max_final = pd.concat([pd.concat(dts_ised_head_max), pd.concat(dts_ised_body_max), pd.concat(dts_ised_body_hotspot_max), pd.concat(dts_ised_hotspot_max), pd.concat(dts_ised_extremity_max)]).reset_index()
            nii_ised_max_final = pd.concat([pd.concat(nii_ised_head_max), pd.concat(nii_ised_body_max), pd.concat(nii_ised_body_hotspot_max), pd.concat(nii_ised_hotspot_max), pd.concat(nii_ised_extremity_max)]).reset_index()
            dss_ised_max_final = pd.concat([pd.concat(dss_ised_head_max), pd.concat(dss_ised_body_max), pd.concat(dss_ised_body_hotspot_max), pd.concat(dss_ised_hotspot_max), pd.concat(dss_ised_extremity_max)]).reset_index()
            
            pce_ised_max_final_filtered = pce_ised_max_final.loc[pce_ised_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            dts_ised_max_final_filtered = dts_ised_max_final.loc[dts_ised_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            nii_ised_max_final_filtered = nii_ised_max_final.loc[nii_ised_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            dss_ised_max_final_filtered = dss_ised_max_final.loc[dss_ised_max_final.groupby(by = ["RF Exposure Condition"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna", "Technology", "Band", "RF Exposure Condition", "Mode", "Power Mode", "Dist (mm)", "Test Position", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
            
            pce_final = pce_ised_max_final_filtered
            dts_final = dts_ised_max_final_filtered
            nii_final = nii_ised_max_final_filtered
            dss_final = dss_ised_max_final_filtered
            
            final_list = [section_1_summary_df, pce_final, dts_final, nii_final, dss_final]
            
            equip_list = ["Section 1 Summary", "PCE", "DTS", "NII", "DSS"]
            
            try:
                abs_filepath = Path(f"{self.s1_ised}").resolve(strict = True)
            
            except FileNotFoundError:
                with pd.ExcelWriter(f"{self.s1_ised}") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(final_list)):  
                        final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
            
            else:
                with pd.ExcelWriter(f"{self.s1_ised}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(final_list)):  
                        final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
        
        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Section 1 ISED Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)