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
            "MSS"       :  ["MSS"],
            "NTN"       :  ["NTN S-Band", "NTN L-Band"]
            }
        self.fcc = {
            "TNE" :    ["LTE B53", "FR1 n53", "MSS"],
            "TNB" :    ["NTN S-Band", "NTN L-Band"],
            "PCE" :    ["GSM 850", "GSM 1900", "W-CDMA B2", "W-CDMA B4", "W-CDMA B5", "LTE B2", "LTE B4",
                        "LTE B5", "LTE B7", "LTE B12", "LTE B13", "LTE B14", "LTE B17", "LTE B25", "LTE B26",
                        "LTE B30", "LTE B38", "LTE B40 (Block A)", "LTE B40 (Block B)", "LTE B41 PC3",
                        "LTE B41 PC2", "LTE B66", "LTE B70", "LTE B71", "LTE B74", "LTE B85", "FR1 n2",
                        "FR1 n5", "FR1 n7", "FR1 n12", "FR1 n13", "FR1 n14", "FR1 n25", "FR1 n26", "FR1 n30",
                        "FR1 n38", "FR1 n40 (Block A)", "FR1 n40 (Block B)", "FR1 n41 PC3", "FR1 n41 PC2", "FR1 n66",
                        "FR1 n70", "FR1 n77 PC3", "FR1 n77 (Block A) PC3", "FR1 n77 (Block B) PC3", "FR1 n77 (Block C) PC3",
                        "FR1 n77 PC2", "FR1 n77 (Block A) PC2", "FR1 n77 (Block B) PC2", "FR1 n77 (Block C) PC2", "FR1 n79 (Narrow) PC3"],
            "CBE" :    ["LTE B48", "FR1 n48"],
            "DTS" :    ["Wi-Fi 2.4 GHz", "802.15.4"],
            "NII" :    ["Wi-Fi 5.2 GHz", "Wi-Fi 5.3 GHz", "Wi-Fi 5.5 GHz", "Wi-Fi 5.8 GHz", "Bluetooth (NB U-NII 1)", "Bluetooth (NB U-NII 3)",
                        "802.15.4ab"],
            "6CD" :    ["U-NII 5", "U-NII 6", "U-NII 7", "U-NII 8"],
            "DSS" :    ["Bluetooth (2.4 GHz)"]
            }
        self.ised = {
            "PCE" :    ["GSM 850", "GSM 1900", "W-CDMA B2", "W-CDMA B4", "W-CDMA B5", "LTE B2", "LTE B4",
                        "LTE B5", "LTE B7", "LTE B12", "LTE B13", "LTE B14", "LTE B17", "LTE B25",
                        "LTE B30", "LTE B38", "LTE B41 PC3", "LTE B41 PC2", "LTE B53", "LTE B66",
                        "LTE B70", "LTE B71", "FR1 n2", "FR1 n5", "FR1 n7", "FR1 n12", "FR1 n13",
                        "FR1 n14", "FR1 n25", "FR1 n30", "FR1 n38", "FR1 n41 PC3", "FR1 n66",
                        "FR1 n41 PC2", "FR1 n53", "FR1 n78 PC3", "FR1 n78 PC2", "MSS", "NTN S-Band", "NTN L-Band"],
            "DTS" :    ["Wi-Fi 2.4 GHz", "802.15.4"],
            "NII" :    ["Wi-Fi 5.2 GHz", "Wi-Fi 5.3 GHz", "Wi-Fi 5.5 GHz", "Wi-Fi 5.8 GHz", "U-NII 5",
                        "Bluetooth (NB U-NII 1)", "Bluetooth (NB U-NII 3)"],
            "DSS" :    ["Bluetooth (2.4 GHz)"],
            "APD" :    ["U-NII 5", "U-NII 6", "U-NII 7", "U-NII 8"]
            }
        
        self.log = log_dir
    
    def section_1_fcc(self):
        
        try:
            cbe_data = [self.data[self.data["Tech"] == self.fcc["CBE"][tech]] for tech in range(len(self.fcc["CBE"]))]
            pce_data = [self.data[self.data["Tech"] == self.fcc["PCE"][tech]] for tech in range(len(self.fcc["PCE"]))]
            tne_data = [self.data[self.data["Tech"] == self.fcc["TNE"][tech]] for tech in range(len(self.fcc["TNE"]))]
            tnb_data = [self.data[self.data["Tech"] == self.fcc["TNB"][tech]] for tech in range(len(self.fcc["TNB"]))]
            dts_data = [self.data[self.data["Tech"] == self.fcc["DTS"][tech]] for tech in range(len(self.fcc["DTS"]))]
            nii_data = [self.data[self.data["Tech"] == self.fcc["NII"][tech]] for tech in range(len(self.fcc["NII"]))]
            dss_data = [self.data[self.data["Tech"] == self.fcc["DSS"][tech]] for tech in range(len(self.fcc["DSS"]))]
            cd6_data = [self.data[self.data["Tech"] == self.fcc["6CD"][tech]] for tech in range(len(self.fcc["6CD"]))]
            
            cbe_separated_data = [cbe_data[no_nan][cbe_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(cbe_data))]
            pce_separated_data = [pce_data[no_nan][pce_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(pce_data))]
            tne_separated_data = [tne_data[no_nan][tne_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(tne_data))]
            tnb_separated_data = [tnb_data[no_nan][tnb_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(tnb_data))]
            dts_separated_data = [dts_data[no_nan][dts_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dts_data))]
            nii_separated_data = [nii_data[no_nan][nii_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(nii_data))]
            dss_separated_data = [dss_data[no_nan][dss_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dss_data))]
            cd6_separated_data = [cd6_data[no_nan][cd6_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(cd6_data))]
            
            fcc_exposure_split(cbe = cbe_separated_data, pce = pce_separated_data, tne = tne_separated_data, tnb = tnb_separated_data, dts = dts_separated_data, nii = nii_separated_data, dss = dss_separated_data, cd6 = cd6_separated_data, directory = self.s1_fcc)
            
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
            pce_data = [self.data[self.data["Tech"] == self.ised["PCE"][tech]] for tech in range(len(self.ised["PCE"]))]
            dts_data = [self.data[self.data["Tech"] == self.ised["DTS"][tech]] for tech in range(len(self.ised["DTS"]))]
            nii_data = [self.data[self.data["Tech"] == self.ised["NII"][tech]] for tech in range(len(self.ised["NII"]))]
            dss_data = [self.data[self.data["Tech"] == self.ised["DSS"][tech]] for tech in range(len(self.ised["DSS"]))]
            apd_data = [self.data[self.data["Tech"] == self.ised["APD"][tech]] for tech in range(len(self.ised["APD"]))]
            
            pce_separated_data = [pce_data[no_nan][pce_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(pce_data))]
            dts_separated_data = [dts_data[no_nan][dts_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dts_data))]
            nii_separated_data = [nii_data[no_nan][nii_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(nii_data))]
            dss_separated_data = [dss_data[no_nan][dss_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(dss_data))]
            apd_separated_data = [apd_data[no_nan][apd_data[no_nan]["1-g Meas. (W/kg)"] != 0] for no_nan in range(len(apd_data))]
            
            ised_exposure_split(pce = pce_separated_data, dts = dts_separated_data, nii = nii_separated_data, dss = dss_separated_data, apd = apd_separated_data, directory = self.s1_ised)
            
        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Section 1 ISED Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)

def fcc_exposure_split(cbe, pce, tne, tnb, dts, nii, dss, cd6, directory):
    
    cbe_head_data = [cbe[head_data][cbe[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(cbe))]
    pce_head_data = [pce[head_data][pce[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(pce))]
    tne_head_data = [tne[head_data][tne[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(tne))]
    tnb_head_data = [tnb[head_data][tnb[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(tnb))]
    dts_head_data = [dts[head_data][dts[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(dts))]
    nii_head_data = [nii[head_data][nii[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(nii))]
    dss_head_data = [dss[head_data][dss[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(dss))]
    cd6_head_data = [cd6[head_data][cd6[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(cd6))]
    
    cbe_body_data = [cbe[body_data][cbe[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(cbe))]
    pce_body_data = [pce[body_data][pce[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(pce))]
    tne_body_data = [tne[body_data][tne[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(tne))]
    tnb_body_data = [tnb[body_data][tnb[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(tnb))]
    dts_body_data = [dts[body_data][dts[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(dts))]
    nii_body_data = [nii[body_data][nii[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(nii))]
    dss_body_data = [dss[body_data][dss[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(dss))]
    cd6_body_data = [cd6[body_data][cd6[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(cd6))]
    
    cbe_body_hotspot_data = [cbe[body_hotspot_data][cbe[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(cbe))]
    pce_body_hotspot_data = [pce[body_hotspot_data][pce[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(pce))]
    tne_body_hotspot_data = [tne[body_hotspot_data][tne[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(tne))]
    tnb_body_hotspot_data = [tnb[body_hotspot_data][tnb[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(tnb))]
    dts_body_hotspot_data = [dts[body_hotspot_data][dts[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(dts))]
    nii_body_hotspot_data = [nii[body_hotspot_data][nii[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(nii))]
    dss_body_hotspot_data = [dss[body_hotspot_data][dss[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(dss))]
    cd6_body_hotspot_data = [cd6[body_hotspot_data][(cd6[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot") & (cd6[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Extremity")] for body_hotspot_data in range(len(cd6))]
    
    cbe_hotspot_data = [cbe[hotspot_data][cbe[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(cbe))]
    pce_hotspot_data = [pce[hotspot_data][pce[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(pce))]
    tne_hotspot_data = [tne[hotspot_data][tne[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(tne))]
    tnb_hotspot_data = [tnb[hotspot_data][tnb[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(tnb))]
    dts_hotspot_data = [dts[hotspot_data][dts[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(dts))]
    nii_hotspot_data = [nii[hotspot_data][nii[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(nii))]
    dss_hotspot_data = [dss[hotspot_data][dss[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(dss))]
    cd6_hotspot_data = [cd6[hotspot_data][cd6[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(cd6))]
    
    cbe_extremity_data = [cbe[extremity_data][cbe[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(cbe))]
    pce_extremity_data = [pce[extremity_data][pce[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(pce))]
    tne_extremity_data = [tne[extremity_data][tne[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(tne))]
    tnb_extremity_data = [tnb[extremity_data][tnb[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(tnb))]
    dts_extremity_data = [dts[extremity_data][dts[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(dts))]
    nii_extremity_data = [nii[extremity_data][nii[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(nii))]
    dss_extremity_data = [dss[extremity_data][dss[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(dss))]
    cd6_extremity_data = [cd6[extremity_data][cd6[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(cd6))]
    
    head          = [cbe_head_data, pce_head_data, tne_head_data, tnb_head_data, dts_head_data, nii_head_data, dss_head_data, cd6_head_data]
    body          = [cbe_body_data, pce_body_data, tne_body_data, tnb_body_data, dts_body_data, nii_body_data, dss_body_data, cd6_body_data]
    bod_hot       = [cbe_body_hotspot_data, pce_body_hotspot_data, tne_body_hotspot_data, tnb_body_hotspot_data, dts_body_hotspot_data, nii_body_hotspot_data, dss_body_hotspot_data, cd6_body_hotspot_data]
    hotspot       = [cbe_hotspot_data, pce_hotspot_data, tne_hotspot_data, tnb_hotspot_data, dts_hotspot_data, nii_hotspot_data, dss_hotspot_data, cd6_hotspot_data]
    extremity     = [cbe_extremity_data, pce_extremity_data, tne_extremity_data, tnb_extremity_data, dts_extremity_data, nii_extremity_data, dss_extremity_data, cd6_extremity_data]
    
    fcc_max_sar_per_exposure(head = head, body = body, bod_hot = bod_hot, hotspot = hotspot, extremity = extremity, directory = directory)

def ised_exposure_split(pce, dts, nii, dss, apd, directory):
    
    pce_head_data = [pce[head_data][pce[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(pce))]
    dts_head_data = [dts[head_data][dts[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(dts))]
    nii_head_data = [nii[head_data][nii[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(nii))]
    dss_head_data = [dss[head_data][dss[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(dss))]
    apd_head_data = [apd[head_data][apd[head_data]["RF Exposure Condition(s)"] == "Head"] for head_data in range(len(apd))]
    
    pce_body_data = [pce[body_data][pce[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(pce))]
    dts_body_data = [dts[body_data][dts[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(dts))]
    nii_body_data = [nii[body_data][nii[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(nii))]
    dss_body_data = [dss[body_data][dss[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(dss))]
    apd_body_data = [apd[body_data][apd[body_data]["RF Exposure Condition(s)"] == "Body-worn"] for body_data in range(len(apd))]
    
    pce_body_hotspot_data = [pce[body_hotspot_data][pce[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(pce))]
    dts_body_hotspot_data = [dts[body_hotspot_data][dts[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(dts))]
    nii_body_hotspot_data = [nii[body_hotspot_data][nii[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(nii))]
    dss_body_hotspot_data = [dss[body_hotspot_data][dss[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(dss))]
    apd_body_hotspot_data = [apd[body_hotspot_data][apd[body_hotspot_data]["RF Exposure Condition(s)"] == "Body & Hotspot"] for body_hotspot_data in range(len(apd))]
    
    pce_hotspot_data = [pce[hotspot_data][pce[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(pce))]
    dts_hotspot_data = [dts[hotspot_data][dts[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(dts))]
    nii_hotspot_data = [nii[hotspot_data][nii[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(nii))]
    dss_hotspot_data = [dss[hotspot_data][dss[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(dss))]
    apd_hotspot_data = [apd[hotspot_data][apd[hotspot_data]["RF Exposure Condition(s)"] == "Hotspot"] for hotspot_data in range(len(apd))]
    
    pce_extremity_data = [pce[extremity_data][pce[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(pce))]
    dts_extremity_data = [dts[extremity_data][dts[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(dts))]
    nii_extremity_data = [nii[extremity_data][nii[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(nii))]
    dss_extremity_data = [dss[extremity_data][dss[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(dss))]
    apd_extremity_data = [apd[extremity_data][apd[extremity_data]["RF Exposure Condition(s)"] == "Extremity"] for extremity_data in range(len(apd))]
    
    pce_extremity_hth_data = [pce[extremity_hth_data][pce[extremity_hth_data]["RF Exposure Condition(s)"] == "Extremity Held-to-Head"] for extremity_hth_data in range(len(pce))]
    dts_extremity_hth_data = [dts[extremity_hth_data][dts[extremity_hth_data]["RF Exposure Condition(s)"] == "Extremity Held-to-Head"] for extremity_hth_data in range(len(dts))]
    nii_extremity_hth_data = [nii[extremity_hth_data][nii[extremity_hth_data]["RF Exposure Condition(s)"] == "Extremity Held-to-Head"] for extremity_hth_data in range(len(nii))]
    dss_extremity_hth_data = [dss[extremity_hth_data][dss[extremity_hth_data]["RF Exposure Condition(s)"] == "Extremity Held-to-Head"] for extremity_hth_data in range(len(dss))]
    apd_extremity_hth_data = [apd[extremity_hth_data][apd[extremity_hth_data]["RF Exposure Condition(s)"] == "Extremity Held-to-Head"] for extremity_hth_data in range(len(apd))]
    
    head          = [pce_head_data, dts_head_data, nii_head_data, dss_head_data, apd_head_data]
    body          = [pce_body_data, dts_body_data, nii_body_data, dss_body_data, apd_body_data]
    bod_hot       = [pce_body_hotspot_data, dts_body_hotspot_data, nii_body_hotspot_data, dss_body_hotspot_data, apd_body_hotspot_data]
    hotspot       = [pce_hotspot_data, dts_hotspot_data, nii_hotspot_data, dss_hotspot_data, apd_hotspot_data]
    extremity     = [pce_extremity_data, dts_extremity_data, nii_extremity_data, dss_extremity_data, apd_extremity_data]
    extremity_hth = [pce_extremity_hth_data, dts_extremity_hth_data, nii_extremity_hth_data, dss_extremity_hth_data, apd_extremity_hth_data]
    
    ised_max_sar_per_exposure(head = head, body = body, bod_hot = bod_hot, hotspot = hotspot, extremity = extremity, extremity_hth = extremity_hth, directory = directory)

def fcc_max_sar_per_exposure(head, body, bod_hot, hotspot, extremity, directory):
    
    cbe = [head[0], body[0], bod_hot[0], hotspot[0], extremity[0]]
    pce = [head[1], body[1], bod_hot[1], hotspot[1], extremity[1]]
    tne = [head[2], body[2], bod_hot[2], hotspot[2], extremity[2]]
    tnb = [head[3], body[3], bod_hot[3], hotspot[3], extremity[3]]
    dts = [head[4], body[4], bod_hot[4], hotspot[4], extremity[4]]
    nii = [head[5], body[5], bod_hot[5], hotspot[5], extremity[5]]
    dss = [head[6], body[6], bod_hot[6], hotspot[6], extremity[6]]
    cd6 = [head[7], body[7], bod_hot[7], hotspot[7], extremity[7]]
    
    cbe_head_max = [cbe[0][max_data][cbe[0][max_data]["1-g Scaled (W/kg)"] == cbe[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cbe[0]))]
    pce_head_max = [pce[0][max_data][pce[0][max_data]["1-g Scaled (W/kg)"] == pce[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[0]))]
    tne_head_max = [tne[0][max_data][tne[0][max_data]["1-g Scaled (W/kg)"] == tne[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tne[0]))]
    tnb_head_max = [tnb[0][max_data][tnb[0][max_data]["1-g Scaled (W/kg)"] == tnb[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tnb[0]))]
    dts_head_max = [dts[0][max_data][dts[0][max_data]["1-g Scaled (W/kg)"] == dts[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[0]))]
    nii_head_max = [nii[0][max_data][nii[0][max_data]["1-g Scaled (W/kg)"] == nii[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[0]))]
    dss_head_max = [dss[0][max_data][dss[0][max_data]["1-g Scaled (W/kg)"] == dss[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[0]))]
    cd6_head_max = [cd6[0][max_data][cd6[0][max_data]["1-g Scaled (W/kg)"] == cd6[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cd6[0]))]
    
    cbe_body_max = [cbe[1][max_data][cbe[1][max_data]["1-g Scaled (W/kg)"] == cbe[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cbe[1]))]
    pce_body_max = [pce[1][max_data][pce[1][max_data]["1-g Scaled (W/kg)"] == pce[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[1]))]
    tne_body_max = [tne[1][max_data][tne[1][max_data]["1-g Scaled (W/kg)"] == tne[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tne[1]))]
    tnb_body_max = [tnb[1][max_data][tnb[1][max_data]["1-g Scaled (W/kg)"] == tnb[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tnb[1]))]
    dts_body_max = [dts[1][max_data][dts[1][max_data]["1-g Scaled (W/kg)"] == dts[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[1]))]
    nii_body_max = [nii[1][max_data][nii[1][max_data]["1-g Scaled (W/kg)"] == nii[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[1]))]
    dss_body_max = [dss[1][max_data][dss[1][max_data]["1-g Scaled (W/kg)"] == dss[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[1]))]
    cd6_body_max = [cd6[1][max_data][cd6[1][max_data]["1-g Scaled (W/kg)"] == cd6[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cd6[1]))]
    
    cbe_boho_max = [cbe[2][max_data][cbe[2][max_data]["1-g Scaled (W/kg)"] == cbe[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cbe[2]))]
    pce_boho_max = [pce[2][max_data][pce[2][max_data]["1-g Scaled (W/kg)"] == pce[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[2]))]
    tne_boho_max = [tne[2][max_data][tne[2][max_data]["1-g Scaled (W/kg)"] == tne[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tne[2]))]
    tnb_boho_max = [tnb[2][max_data][tnb[2][max_data]["1-g Scaled (W/kg)"] == tnb[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tnb[2]))]
    dts_boho_max = [dts[2][max_data][dts[2][max_data]["1-g Scaled (W/kg)"] == dts[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[2]))]
    nii_boho_max = [nii[2][max_data][nii[2][max_data]["1-g Scaled (W/kg)"] == nii[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[2]))]
    dss_boho_max = [dss[2][max_data][dss[2][max_data]["1-g Scaled (W/kg)"] == dss[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[2]))]
    cd6_boho_max = [cd6[2][max_data][cd6[2][max_data]["1-g Scaled (W/kg)"] == cd6[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cd6[2]))]
    
    cbe_hoty_max = [cbe[3][max_data][cbe[3][max_data]["1-g Scaled (W/kg)"] == cbe[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cbe[3]))]
    pce_hoty_max = [pce[3][max_data][pce[3][max_data]["1-g Scaled (W/kg)"] == pce[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[3]))]
    tne_hoty_max = [tne[3][max_data][tne[3][max_data]["1-g Scaled (W/kg)"] == tne[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tne[3]))]
    tnb_hoty_max = [tnb[3][max_data][tnb[3][max_data]["1-g Scaled (W/kg)"] == tnb[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(tnb[3]))]
    dts_hoty_max = [dts[3][max_data][dts[3][max_data]["1-g Scaled (W/kg)"] == dts[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[3]))]
    nii_hoty_max = [nii[3][max_data][nii[3][max_data]["1-g Scaled (W/kg)"] == nii[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[3]))]
    dss_hoty_max = [dss[3][max_data][dss[3][max_data]["1-g Scaled (W/kg)"] == dss[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[3]))]
    cd6_hoty_max = [cd6[3][max_data][cd6[3][max_data]["1-g Scaled (W/kg)"] == cd6[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(cd6[3]))]
    
    cbe_extr_max = [cbe[4][max_data][cbe[4][max_data]["10-g Scaled (W/kg)"] == cbe[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(cbe[4]))]
    pce_extr_max = [pce[4][max_data][pce[4][max_data]["10-g Scaled (W/kg)"] == pce[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(pce[4]))]
    tne_extr_max = [tne[4][max_data][tne[4][max_data]["10-g Scaled (W/kg)"] == tne[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(tne[4]))]
    tnb_extr_max = [tnb[4][max_data][tnb[4][max_data]["10-g Scaled (W/kg)"] == tnb[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(tnb[4]))]
    dts_extr_max = [dts[4][max_data][dts[4][max_data]["10-g Scaled (W/kg)"] == dts[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dts[4]))]
    nii_extr_max = [nii[4][max_data][nii[4][max_data]["10-g Scaled (W/kg)"] == nii[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(nii[4]))]
    dss_extr_max = [dss[4][max_data][dss[4][max_data]["10-g Scaled (W/kg)"] == dss[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dss[4]))]
    cd6_extr_max = [cd6[4][max_data][cd6[4][max_data]["10-g Scaled (W/kg)"] == cd6[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(cd6[4]))]
    
    cbe_sec_1 = [pd.concat(cbe_head_max), pd.concat(cbe_body_max), pd.concat(cbe_boho_max), pd.concat(cbe_hoty_max), pd.concat(cbe_extr_max)]
    pce_sec_1 = [pd.concat(pce_head_max), pd.concat(pce_body_max), pd.concat(pce_boho_max), pd.concat(pce_hoty_max), pd.concat(pce_extr_max)]
    tne_sec_1 = [pd.concat(tne_head_max), pd.concat(tne_body_max), pd.concat(tne_boho_max), pd.concat(tne_hoty_max), pd.concat(tne_extr_max)]
    tnb_sec_1 = [pd.concat(tnb_head_max), pd.concat(tnb_body_max), pd.concat(tnb_boho_max), pd.concat(tnb_hoty_max), pd.concat(tnb_extr_max)]
    dts_sec_1 = [pd.concat(dts_head_max), pd.concat(dts_body_max), pd.concat(dts_boho_max), pd.concat(dts_hoty_max), pd.concat(dts_extr_max)]
    nii_sec_1 = [pd.concat(nii_head_max), pd.concat(nii_body_max), pd.concat(nii_boho_max), pd.concat(nii_hoty_max), pd.concat(nii_extr_max)]
    dss_sec_1 = [pd.concat(dss_head_max), pd.concat(dss_body_max), pd.concat(dss_boho_max), pd.concat(dss_hoty_max), pd.concat(dss_extr_max)]
    cd6_sec_1 = [pd.concat(cd6_head_max), pd.concat(cd6_body_max), pd.concat(cd6_boho_max), pd.concat(cd6_hoty_max), pd.concat(cd6_extr_max)]
    
    build_fcc_sec_1_df(cbe = cbe_sec_1, pce = pce_sec_1, tne = tne_sec_1, tnb = tnb_sec_1, dts = dts_sec_1, nii = nii_sec_1, dss = dss_sec_1, cd6 = cd6_sec_1, directory = directory)

def ised_max_sar_per_exposure(head, body, bod_hot, hotspot, extremity, extremity_hth, directory):
    
    pce = [head[0], body[0], bod_hot[0], hotspot[0], extremity[0], extremity_hth[0]]
    dts = [head[1], body[1], bod_hot[1], hotspot[1], extremity[1], extremity_hth[1]]
    nii = [head[2], body[2], bod_hot[2], hotspot[2], extremity[2], extremity_hth[2]]
    dss = [head[3], body[3], bod_hot[3], hotspot[3], extremity[3], extremity_hth[3]]
    apd = [head[4], body[4], bod_hot[4], hotspot[4], extremity[4], extremity_hth[4]]
    
    pce_head_max = [pce[0][max_data][pce[0][max_data]["1-g Scaled (W/kg)"] == pce[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[0]))]
    dts_head_max = [dts[0][max_data][dts[0][max_data]["1-g Scaled (W/kg)"] == dts[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[0]))]
    nii_head_max = [nii[0][max_data][nii[0][max_data]["1-g Scaled (W/kg)"] == nii[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[0]))]
    dss_head_max = [dss[0][max_data][dss[0][max_data]["1-g Scaled (W/kg)"] == dss[0][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[0]))]
    apd_head_max = [apd[0][max_data][apd[0][max_data]["APD Scaled (W/m2)"] == apd[0][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[0]))]
    
    pce_body_max = [pce[1][max_data][pce[1][max_data]["1-g Scaled (W/kg)"] == pce[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[1]))]
    dts_body_max = [dts[1][max_data][dts[1][max_data]["1-g Scaled (W/kg)"] == dts[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[1]))]
    nii_body_max = [nii[1][max_data][nii[1][max_data]["1-g Scaled (W/kg)"] == nii[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[1]))]
    dss_body_max = [dss[1][max_data][dss[1][max_data]["1-g Scaled (W/kg)"] == dss[1][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[1]))]
    apd_body_max = [apd[1][max_data][apd[1][max_data]["APD Scaled (W/m2)"] == apd[1][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[1]))]
    
    pce_boho_max = [pce[2][max_data][pce[2][max_data]["1-g Scaled (W/kg)"] == pce[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[2]))]
    dts_boho_max = [dts[2][max_data][dts[2][max_data]["1-g Scaled (W/kg)"] == dts[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[2]))]
    nii_boho_max = [nii[2][max_data][nii[2][max_data]["1-g Scaled (W/kg)"] == nii[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[2]))]
    dss_boho_max = [dss[2][max_data][dss[2][max_data]["1-g Scaled (W/kg)"] == dss[2][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[2]))]
    apd_boho_max = [apd[2][max_data][apd[2][max_data]["APD Scaled (W/m2)"] == apd[2][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[2]))]
    
    pce_hoty_max = [pce[3][max_data][pce[3][max_data]["1-g Scaled (W/kg)"] == pce[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(pce[3]))]
    dts_hoty_max = [dts[3][max_data][dts[3][max_data]["1-g Scaled (W/kg)"] == dts[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dts[3]))]
    nii_hoty_max = [nii[3][max_data][nii[3][max_data]["1-g Scaled (W/kg)"] == nii[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(nii[3]))]
    dss_hoty_max = [dss[3][max_data][dss[3][max_data]["1-g Scaled (W/kg)"] == dss[3][max_data]["1-g Scaled (W/kg)"].max()] for max_data in range(len(dss[3]))]
    apd_hoty_max = [apd[3][max_data][apd[3][max_data]["APD Scaled (W/m2)"] == apd[3][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[3]))]
    
    pce_extr_max = [pce[4][max_data][pce[4][max_data]["10-g Scaled (W/kg)"] == pce[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(pce[4]))]
    dts_extr_max = [dts[4][max_data][dts[4][max_data]["10-g Scaled (W/kg)"] == dts[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dts[4]))]
    nii_extr_max = [nii[4][max_data][nii[4][max_data]["10-g Scaled (W/kg)"] == nii[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(nii[4]))]
    dss_extr_max = [dss[4][max_data][dss[4][max_data]["10-g Scaled (W/kg)"] == dss[4][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dss[4]))]
    apd_extr_max = [apd[4][max_data][apd[4][max_data]["APD Scaled (W/m2)"] == apd[4][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[4]))]
    
    pce_ehth_max = [pce[5][max_data][pce[5][max_data]["10-g Scaled (W/kg)"] == pce[5][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(pce[5]))]
    dts_ehth_max = [dts[5][max_data][dts[5][max_data]["10-g Scaled (W/kg)"] == dts[5][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dts[5]))]
    nii_ehth_max = [nii[5][max_data][nii[5][max_data]["10-g Scaled (W/kg)"] == nii[5][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(nii[5]))]
    dss_ehth_max = [dss[5][max_data][dss[5][max_data]["10-g Scaled (W/kg)"] == dss[5][max_data]["10-g Scaled (W/kg)"].max()] for max_data in range(len(dss[5]))]
    apd_ehth_max = [apd[5][max_data][apd[5][max_data]["APD Scaled (W/m2)"] == apd[5][max_data]["APD Scaled (W/m2)"].max()] for max_data in range(len(apd[5]))]
    
    pce_sec_1 = [pd.concat(pce_head_max), pd.concat(pce_body_max), pd.concat(pce_boho_max), pd.concat(pce_hoty_max), pd.concat(pce_extr_max), pd.concat(pce_ehth_max)]
    dts_sec_1 = [pd.concat(dts_head_max), pd.concat(dts_body_max), pd.concat(dts_boho_max), pd.concat(dts_hoty_max), pd.concat(dts_extr_max), pd.concat(dts_ehth_max)]
    nii_sec_1 = [pd.concat(nii_head_max), pd.concat(nii_body_max), pd.concat(nii_boho_max), pd.concat(nii_hoty_max), pd.concat(nii_extr_max), pd.concat(nii_ehth_max)]
    dss_sec_1 = [pd.concat(dss_head_max), pd.concat(dss_body_max), pd.concat(dss_boho_max), pd.concat(dss_hoty_max), pd.concat(dss_extr_max), pd.concat(dss_ehth_max)]
    apd_sec_1 = [pd.concat(apd_head_max), pd.concat(apd_body_max), pd.concat(apd_boho_max), pd.concat(apd_hoty_max), pd.concat(apd_extr_max), pd.concat(apd_ehth_max)]
    
    build_ised_sec_1_df(pce = pce_sec_1, dts = dts_sec_1, nii = nii_sec_1, dss = dss_sec_1, apd = apd_sec_1, directory = directory)

def build_fcc_sec_1_df(cbe, pce, tne, tnb, dts, nii, dss, cd6, directory):
    
    if cbe[0].size == 0:
        cbe_head = "N/A"
    else:
        cbe_head_max = cbe[0]["1-g Scaled (W/kg)"].max()
        cbe_head = float(cbe_head_max)
    if cbe[1].size == 0:
        cbe_body = "N/A"
    else:
        cbe_body_max = cbe[1]["1-g Scaled (W/kg)"].max()
        cbe_body = float(cbe_body_max)
    if cbe[2].size == 0:
        cbe_boho = "N/A"
    else:
        cbe_boho_max = cbe[2]["1-g Scaled (W/kg)"].max()
        cbe_boho = float(cbe_boho_max)
    if cbe[3].size == 0:
        cbe_hoty = "N/A"
    else:
        cbe_hoty_max = cbe[3]["1-g Scaled (W/kg)"].max()
        cbe_hoty = float(cbe_hoty_max)
    if cbe[4].size == 0:
        cbe_extr = "N/A"
    else:
        cbe_extr_max = cbe[4]["1-g Scaled (W/kg)"].max()
        cbe_extr = float(cbe_extr_max)
    
    if pce[0].size == 0:
        pce_head = "N/A"
    else:
        pce_head_max = pce[0]["1-g Scaled (W/kg)"].max()
        pce_head = float(pce_head_max)
    if pce[1].size == 0:
        pce_body = "N/A"
    else:
        pce_body_max = pce[1]["1-g Scaled (W/kg)"].max()
        pce_body = float(pce_body_max)
    if pce[2].size == 0:
        pce_boho = "N/A"
    else:
        pce_boho_max = pce[2]["1-g Scaled (W/kg)"].max()
        pce_boho = float(pce_boho_max)
    if pce[3].size == 0:
        pce_hoty = "N/A"
    else:
        pce_hoty_max = pce[3]["1-g Scaled (W/kg)"].max()
        pce_hoty = float(pce_hoty_max)
    if pce[4].size == 0:
        pce_extr = "N/A"
    else:
        pce_extr_max = pce[4]["10-g Scaled (W/kg)"].max()
        pce_extr = float(pce_extr_max)
    
    if tne[0].size == 0:
        tne_head = "N/A"
    else:
        tne_head_max = tne[0]["1-g Scaled (W/kg)"].max()
        tne_head = float(tne_head_max)
    if tne[1].size == 0:
        tne_body = "N/A"
    else:
        tne_body_max = tne[1]["1-g Scaled (W/kg)"].max()
        tne_body = float(tne_body_max)
    if tne[2].size == 0:
        tne_boho = "N/A"
    else:
        tne_boho_max = tne[2]["1-g Scaled (W/kg)"].max()
        tne_boho = float(tne_boho_max)
    if tne[3].size == 0:
        tne_hoty = "N/A"
    else:
        tne_hoty_max = tne[3]["1-g Scaled (W/kg)"].max()
        tne_hoty = float(tne_hoty_max)
    if tne[4].size == 0:
        tne_extr = "N/A"
    else:
        tne_extr_max = tne[4]["10-g Scaled (W/kg)"].max()
        tne_extr = float(tne_extr_max)
    
    if tnb[0].size == 0:
        tnb_head = "N/A"
    else:
        tnb_head_max = tnb[0]["1-g Scaled (W/kg)"].max()
        tnb_head = float(tnb_head_max)
    if tnb[1].size == 0:
        tnb_body = "N/A"
    else:
        tnb_body_max = tnb[1]["1-g Scaled (W/kg)"].max()
        tnb_body = float(tnb_body_max)
    if tnb[2].size == 0:
        tnb_boho = "N/A"
    else:
        tnb_boho_max = tnb[2]["1-g Scaled (W/kg)"].max()
        tnb_boho = float(tnb_boho_max)
    if tnb[3].size == 0:
        tnb_hoty = "N/A"
    else:
        tnb_hoty_max = tnb[3]["1-g Scaled (W/kg)"].max()
        tnb_hoty = float(tnb_hoty_max)
    if tnb[4].size == 0:
        tnb_extr = "N/A"
    else:
        tnb_extr_max = tnb[4]["10-g Scaled (W/kg)"].max()
        tnb_extr = float(tnb_extr_max)
    
    if dts[0].size == 0:
        dts_head = "N/A"
    else:
        dts_head_max = dts[0]["1-g Scaled (W/kg)"].max()
        dts_head = float(dts_head_max)
    if dts[1].size == 0:
        dts_body = "N/A"
    else:
        dts_body_max = dts[1]["1-g Scaled (W/kg)"].max()
        dts_body = float(dts_body_max)
    if dts[2].size == 0:
        dts_boho = "N/A"
    else:
        dts_boho_max = dts[2]["1-g Scaled (W/kg)"].max()
        dts_boho = float(dts_boho_max)
    if dts[3].size == 0:
        dts_hoty = "N/A"
    else:
        dts_hoty_max = dts[3]["1-g Scaled (W/kg)"].max()
        dts_hoty = float(dts_hoty_max)
    if dts[4].size == 0:
        dts_extr = "N/A"
    else:
        dts_extr_max = dts[4]["10-g Scaled (W/kg)"].max()
        dts_extr = float(dts_extr_max)
    
    if nii[0].size == 0:
        nii_head = "N/A"
    else:
        nii_head_max = nii[0]["1-g Scaled (W/kg)"].max()
        nii_head = float(nii_head_max)
    if nii[1].size == 0:
        nii_body = "N/A"
    else:
        nii_body_max = nii[1]["1-g Scaled (W/kg)"].max()
        nii_body = float(nii_body_max)
    if nii[2].size == 0:
        nii_boho = "N/A"
    else:
        nii_boho_max = nii[2]["1-g Scaled (W/kg)"].max()
        nii_boho = float(nii_boho_max)
    if nii[3].size == 0:
        nii_hoty = "N/A"
    else:
        nii_hoty_max = nii[3]["1-g Scaled (W/kg)"].max()
        nii_hoty = float(nii_hoty_max)
    if nii[4].size == 0:
        nii_extr = "N/A"
    else:
        nii_extr_max = nii[4]["10-g Scaled (W/kg)"].max()
        nii_extr = float(nii_extr_max)
    
    if dss[0].size == 0:
        dss_head = "N/A"
    else:
        dss_head_max = dss[0]["1-g Scaled (W/kg)"].max()
        dss_head = float(dss_head_max)
    if dss[1].size == 0:
        dss_body = "N/A"
    else:
        dss_body_max = dss[1]["1-g Scaled (W/kg)"].max()
        dss_body = float(dss_body_max)
    if dss[2].size == 0:
        dss_boho = "N/A"
    else:
        dss_boho_max = dss[2]["1-g Scaled (W/kg)"].max()
        dss_boho = float(dss_boho_max)
    if dss[3].size == 0:
        dss_hoty = "N/A"
    else:
        dss_hoty_max = dss[3]["1-g Scaled (W/kg)"].max()
        dss_hoty = float(dss_hoty_max)
    if dss[4].size == 0:
        dss_extr = "N/A"
    else:
        dss_extr_max = dss[4]["10-g Scaled (W/kg)"].max()
        dss_extr = float(dss_extr_max)
    
    if cd6[0].size == 0:
        cd6_head = "N/A"
    else:
        cd6_head_max = cd6[0]["1-g Scaled (W/kg)"].max()
        cd6_head = float(cd6_head_max)
    if cd6[1].size == 0:
        cd6_body = "N/A"
    else:
        cd6_body_max = cd6[1]["1-g Scaled (W/kg)"].max()
        cd6_body = float(cd6_body_max)
    if cd6[2].size == 0:
        cd6_boho = "N/A"
    else:
        cd6_boho_max = cd6[2]["1-g Scaled (W/kg)"].max()
        cd6_boho = float(cd6_boho_max)
    if cd6[3].size == 0:
        cd6_hoty = "N/A"
    else:
        cd6_hoty_max = cd6[3]["1-g Scaled (W/kg)"].max()
        cd6_hoty = float(cd6_hoty_max)
    if cd6[4].size == 0:
        cd6_extr = "N/A"
    else:
        cd6_extr_max = cd6[4]["10-g Scaled (W/kg)"].max()
        cd6_extr = float(cd6_extr_max)
    
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
            tne_boho,
            tne_hoty,
            tne_extr
            ],
        "TNB": [
            tnb_head,
            tnb_body,
            tnb_boho,
            tnb_hoty,
            tnb_extr
            ],
        "PCE": [
            pce_head,
            pce_body,
            pce_boho,
            pce_hoty,
            pce_extr
            ],
        "CBE": [
            cbe_head,
            cbe_body,
            cbe_boho,
            cbe_hoty,
            cbe_extr
            ],
        "DTS": [
            dts_head,
            dts_body,
            dts_boho,
            dts_hoty,
            dts_extr
            ],
        "NII": [
            nii_head,
            nii_body,
            nii_boho,
            nii_hoty,
            nii_extr
            ],
        "6CD": [
            cd6_head,
            cd6_body,
            cd6_boho,
            cd6_hoty,
            cd6_extr
            ],
        "DSS": [
            dss_head,
            dss_body,
            dss_boho,
            dss_hoty,
            dss_extr
            ],
        }
    
    section_1_summary_df = pd.DataFrame(section_1_summary)
    
    cbe_max_final = pd.concat([cbe[0], cbe[1], cbe[2], cbe[3], cbe[4]]).reset_index()
    pce_max_final = pd.concat([pce[0], pce[1], pce[2], pce[3], pce[4]]).reset_index()
    tne_max_final = pd.concat([tne[0], tne[1], tne[2], tne[3], tne[4]]).reset_index()
    tnb_max_final = pd.concat([tnb[0], tnb[1], tnb[2], tnb[3], tnb[4]]).reset_index()
    dts_max_final = pd.concat([dts[0], dts[1], dts[2], dts[3], dts[4]]).reset_index()
    nii_max_final = pd.concat([nii[0], nii[1], nii[2], nii[3], nii[4]]).reset_index()
    dss_max_final = pd.concat([dss[0], dss[1], dss[2], dss[3], dss[4]]).reset_index()
    cd6_max_final = pd.concat([cd6[0], cd6[1], cd6[2], cd6[3], cd6[4]]).reset_index()
    
    tne_max_final_filtered = tne_max_final.loc[tne_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    tnb_max_final_filtered = tnb_max_final.loc[tnb_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    pce_max_final_filtered = pce_max_final.loc[pce_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    cbe_max_final_filtered = cbe_max_final.loc[cbe_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    dts_max_final_filtered = dts_max_final.loc[dts_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    nii_max_final_filtered = nii_max_final.loc[nii_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    cd6_max_final_filtered = cd6_max_final.loc[cd6_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    dss_max_final_filtered = dss_max_final.loc[dss_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    tne_final = tne_max_final_filtered
    tnb_final = tnb_max_final_filtered
    pce_final = pce_max_final_filtered
    cbe_final = cbe_max_final_filtered
    dts_final = dts_max_final_filtered
    nii_final = nii_max_final_filtered
    cd6_final = cd6_max_final_filtered
    dss_final = dss_max_final_filtered
    
    final_list = [section_1_summary_df, tne_final, tnb_final, pce_final, cbe_final, dts_final, nii_final, cd6_final, dss_final]
    
    equip_list = ["Section 1 Summary", "TNE", "TNB", "PCE", "CBE", "DTS", "NII", "6CD", "DSS"]
    
    try:
        abs_filepath = Path(f"{directory}").resolve(strict = True)
    
    except FileNotFoundError:
        with pd.ExcelWriter(f"{directory}") as writer: # pylint: disable=abstract-class-instantiated
            for nonsense in range(len(final_list)):  
                final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
    
    else:
        with pd.ExcelWriter(f"{directory}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
            for nonsense in range(len(final_list)):  
                final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)

def build_ised_sec_1_df(pce, dts, nii, dss, apd, directory):
    
    if pce[0].size == 0:
        pce_head = "N/A"
    else:
        pce_head_max = pce[0]["1-g Scaled (W/kg)"].max()
        pce_head = float(pce_head_max)
    if pce[1].size == 0:
        pce_body = "N/A"
    else:
        pce_body_max = pce[1]["1-g Scaled (W/kg)"].max()
        pce_body = float(pce_body_max)
    if pce[2].size == 0:
        pce_boho = "N/A"
    else:
        pce_boho_max = pce[2]["1-g Scaled (W/kg)"].max()
        pce_boho = float(pce_boho_max)
    if pce[3].size == 0:
        pce_hoty = "N/A"
    else:
        pce_hoty_max = pce[3]["1-g Scaled (W/kg)"].max()
        pce_hoty = float(pce_hoty_max)
    if pce[4].size == 0:
        pce_extr = "N/A"
    else:
        pce_extr_max = pce[4]["10-g Scaled (W/kg)"].max()
        pce_extr = float(pce_extr_max)
    if pce[5].size == 0:
        pce_ehth = "N/A"
    else:
        pce_ehth_max = pce[5]["10-g Scaled (W/kg)"].max()
        pce_ehth = float(pce_ehth_max)
    
    if dts[0].size == 0:
        dts_head = "N/A"
    else:
        dts_head_max = dts[0]["1-g Scaled (W/kg)"].max()
        dts_head = float(dts_head_max)
    if dts[1].size == 0:
        dts_body = "N/A"
    else:
        dts_body_max = dts[1]["1-g Scaled (W/kg)"].max()
        dts_body = float(dts_body_max)
    if dts[2].size == 0:
        dts_boho = "N/A"
    else:
        dts_boho_max = dts[2]["1-g Scaled (W/kg)"].max()
        dts_boho = float(dts_boho_max)
    if dts[3].size == 0:
        dts_hoty = "N/A"
    else:
        dts_hoty_max = dts[3]["1-g Scaled (W/kg)"].max()
        dts_hoty = float(dts_hoty_max)
    if dts[4].size == 0:
        dts_extr = "N/A"
    else:
        dts_extr_max = dts[4]["10-g Scaled (W/kg)"].max()
        dts_extr = float(dts_extr_max)
    if dts[5].size == 0:
        dts_ehth = "N/A"
    else:
        dts_ehth_max = dts[5]["10-g Scaled (W/kg)"].max()
        dts_ehth = float(dts_ehth_max)
    
    if nii[0].size == 0:
        nii_head = "N/A"
    else:
        nii_head_max = nii[0]["1-g Scaled (W/kg)"].max()
        nii_head = float(nii_head_max)
    if nii[1].size == 0:
        nii_body = "N/A"
    else:
        nii_body_max = nii[1]["1-g Scaled (W/kg)"].max()
        nii_body = float(nii_body_max)
    if nii[2].size == 0:
        nii_boho = "N/A"
    else:
        nii_boho_max = nii[2]["1-g Scaled (W/kg)"].max()
        nii_boho = float(nii_boho_max)
    if nii[3].size == 0:
        nii_hoty = "N/A"
    else:
        nii_hoty_max = nii[3]["1-g Scaled (W/kg)"].max()
        nii_hoty = float(nii_hoty_max)
    if nii[4].size == 0:
        nii_extr = "N/A"
    else:
        nii_extr_max = nii[4]["10-g Scaled (W/kg)"].max()
        nii_extr = float(nii_extr_max)
    if nii[5].size == 0:
        nii_ehth = "N/A"
    else:
        nii_ehth_max = nii[5]["10-g Scaled (W/kg)"].max()
        nii_ehth = float(nii_ehth_max)
    
    if dss[0].size == 0:
        dss_head = "N/A"
    else:
        dss_head_max = dss[0]["1-g Scaled (W/kg)"].max()
        dss_head = float(dss_head_max)
    if dss[1].size == 0:
        dss_body = "N/A"
    else:
        dss_body_max = dss[1]["1-g Scaled (W/kg)"].max()
        dss_body = float(dss_body_max)
    if dss[2].size == 0:
        dss_boho = "N/A"
    else:
        dss_boho_max = dss[2]["1-g Scaled (W/kg)"].max()
        dss_boho = float(dss_boho_max)
    if dss[3].size == 0:
        dss_hoty = "N/A"
    else:
        dss_hoty_max = dss[3]["1-g Scaled (W/kg)"].max()
        dss_hoty = float(dss_hoty_max)
    if dss[4].size == 0:
        dss_extr = "N/A"
    else:
        dss_extr_max = dss[4]["10-g Scaled (W/kg)"].max()
        dss_extr = float(dss_extr_max)
    if dss[5].size == 0:
        dss_ehth = "N/A"
    else:
        dss_ehth_max = dss[5]["10-g Scaled (W/kg)"].max()
        dss_ehth = float(dss_ehth_max)
    
    if apd[0].size == 0:
        apd_head = "N/A"
    else:
        apd_head_max = apd[0]["1-g Scaled (W/kg)"].max()
        apd_head = float(apd_head_max)
    if apd[1].size == 0:
        apd_body = "N/A"
    else:
        apd_body_max = apd[1]["1-g Scaled (W/kg)"].max()
        apd_body = float(apd_body_max)
    if apd[2].size == 0:
        apd_boho = "N/A"
    else:
        apd_boho_max = apd[2]["1-g Scaled (W/kg)"].max()
        apd_boho = float(apd_boho_max)
    if apd[3].size == 0:
        apd_hoty = "N/A"
    else:
        apd_hoty_max = apd[3]["1-g Scaled (W/kg)"].max()
        apd_hoty = float(apd_hoty_max)
    if apd[4].size == 0:
        apd_extr = "N/A"
    else:
        apd_extr_max = apd[4]["10-g Scaled (W/kg)"].max()
        apd_extr = float(apd_extr_max)
    if apd[5].size == 0:
        apd_ehth = "N/A"
    else:
        apd_ehth_max = apd[5]["10-g Scaled (W/kg)"].max()
        apd_ehth = float(apd_ehth_max)
    
    section_1_summary = {
        "RF Exposure Condition": [
            "Head",
            "Body-worn",
            "Body & Hotspot",
            "Hotspot",
            "Extremity",
            "Extremity Held-to-Head"
            ],
        "PCE": [
            pce_head,
            pce_body,
            pce_boho,
            pce_hoty,
            pce_extr,
            pce_ehth
            ],
        "DTS": [
            dts_head,
            dts_body,
            dts_boho,
            dts_hoty,
            dts_extr,
            dts_ehth
            ],
        "NII": [
            nii_head,
            nii_body,
            nii_boho,
            nii_hoty,
            nii_extr,
            nii_ehth
            ],
        "DSS": [
            dss_head,
            dss_body,
            dss_boho,
            dss_hoty,
            dss_extr,
            dss_ehth
            ],
        "APD": [
            apd_head,
            apd_body,
            apd_boho,
            apd_hoty,
            apd_extr,
            apd_ehth
            ]
        }
    
    section_1_summary_df = pd.DataFrame(section_1_summary)
    
    pce_max_final = pd.concat([pce[0], pce[1], pce[2], pce[3], pce[4], pce[5]]).reset_index()
    dts_max_final = pd.concat([dts[0], dts[1], dts[2], dts[3], dts[4], dts[5]]).reset_index()
    nii_max_final = pd.concat([nii[0], nii[1], nii[2], nii[3], nii[4], nii[5]]).reset_index()
    dss_max_final = pd.concat([dss[0], dss[1], dss[2], dss[3], dss[4], dss[5]]).reset_index()
    apd_max_final = pd.concat([apd[0], apd[1], apd[2], apd[3], apd[4], apd[5]]).reset_index()
    
    pce_max_final_filtered = pce_max_final.loc[pce_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    dts_max_final_filtered = dts_max_final.loc[dts_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    nii_max_final_filtered = nii_max_final.loc[nii_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    dss_max_final_filtered = dss_max_final.loc[dss_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    apd_max_final_filtered = apd_max_final.loc[apd_max_final.groupby(by = ["RF Exposure Condition(s)"])["1-g Scaled (W/kg)"].idxmax()].sort_index().filter(items = ["Antenna(s)", "Technology", "Band", "RF Exposure Condition(s)", "Mode(s)", "Power Mode(s)", "Dist (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"])
    
    pce_final = pce_max_final_filtered
    dts_final = dts_max_final_filtered
    nii_final = nii_max_final_filtered
    dss_final = dss_max_final_filtered
    apd_final = apd_max_final_filtered
    
    final_list = [section_1_summary_df, pce_final, dts_final, nii_final, dss_final, apd_final]
    
    equip_list = ["Section 1 Summary", "PCE", "DTS", "NII", "DSS", "APD"]
    
    try:
        abs_filepath = Path(f"{directory}").resolve(strict = True)
    
    except FileNotFoundError:
        with pd.ExcelWriter(f"{directory}") as writer: # pylint: disable=abstract-class-instantiated
            for nonsense in range(len(final_list)):  
                final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)
    
    else:
        with pd.ExcelWriter(f"{directory}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
            for nonsense in range(len(final_list)):  
                final_list[nonsense].to_excel(writer, sheet_name = f"{equip_list[nonsense]}", index=False)