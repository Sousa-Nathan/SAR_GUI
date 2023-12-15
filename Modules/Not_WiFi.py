""" Cleans up the raw data tables for the report and generates a summary of worst-case SAR """

import pandas as pd
import logging
import os

from pathlib import Path

class Not_WiFi:

    def __init__(self, data: list, tech_list: list, transmitter_names: list, exposure_conditions: list, reported_results_filepath: str, summary_results_filepath: str, smtx_results_filepath: str, log_dir: str):
        self.data = pd.concat(data)
        self.techlist = tech_list
        self.expose = exposure_conditions
        self.trans = transmitter_names
        self.rrf = reported_results_filepath
        self.srf = summary_results_filepath
        self.stx = smtx_results_filepath
        self.log = log_dir
    
    def reported_tech_results(self):
        try:
            reported_tech_results = [self.data[self.data["Tech"] == self.techlist[tech]] for tech in range(len(self.techlist))]
            
            reported_tech_results_filter = [reported_tech_results[tech].filter(items = ["Antenna(s)", "RF Exposure Condition", "Mode(s)", "Power Mode(s)", "Dist. (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "Area Scan Max. SAR (W/kg)","TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)", "Peak SAR (x, y, z)", "Power Drift (dB)", "Plot No."]) for tech in range(len(reported_tech_results))]
            
            drop_zero = [reported_tech_results_filter[zero][reported_tech_results_filter[zero]["1-g Meas. (W/kg)"] != 0] for zero in range(len(reported_tech_results_filter))]
            
            new_reported_tech_results = [drop_zero[blank].dropna(subset = ["1-g Meas. (W/kg)"]) for blank in range(len(drop_zero))]
            
            stupid_eagle_request = [new_reported_tech_results[col_name].rename(columns = {"TuP Limit (dBm)": "Max Output Pwr (dBm)"}) for col_name in range(len(new_reported_tech_results))]
            
            try:
                abs_filepath_rrf = Path(f"{self.rrf}").resolve(strict = True)
            
            except FileNotFoundError:
                with pd.ExcelWriter(f"{self.rrf}") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(stupid_eagle_request)):  
                        stupid_eagle_request[nonsense].to_excel(writer, sheet_name = f"{self.techlist[nonsense]}", index = False)
            
            else:
                with pd.ExcelWriter(f"{self.rrf}", mode = "a") as writer: # pylint: disable=abstract-class-instantiated
                    for nonsense in range(len(stupid_eagle_request)):  
                        stupid_eagle_request[nonsense].to_excel(writer, sheet_name = f"{self.techlist[nonsense]}", index=False)
            
            return new_reported_tech_results
        
        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Not Wi-Fi Reported Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)
    
    def summary_tech_results(self):
        try:
            summary_data = self.data
            summary_data["1-g Meas. (W/kg)"] = summary_data["1-g Meas. (W/kg)"].fillna(0)
            non_zero_data = summary_data[summary_data["1-g Meas. (W/kg)"] != 0]
            
            trans_list = [non_zero_data[non_zero_data["Antenna(s)"] == self.trans[index]] for index, name in enumerate(self.trans)]
            
            summary_results = []
            for data in range(len(trans_list)):
                for tech in range(len(self.techlist)):
                    ant_tech_filter = trans_list[data][trans_list[data]["Tech"] == self.techlist[tech]]
                    summary_results.append(ant_tech_filter.loc[ant_tech_filter.groupby(by = "RF Exposure Condition")["1-g Scaled (W/kg)"].idxmax()].sort_index())
            
            new_summary_results = [summary_results[tech].filter(items = ["System Check Date", "Test Date", "Checked By", "Lab Location", "Sample No.", "Antenna(s)", "Technology", "Band", "RF Exposure Condition", "Mode(s)", "Power Mode(s)", "Dist. (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "Duty Cycle (%)", "Area Scan Max. SAR (W/kg)", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"]) for tech in range(len(summary_results))]
            
            final_summary = pd.concat(new_summary_results)
            
            try:
                abs_summary_filepath = Path(f"{self.srf}").resolve(strict = True)
            
            except FileNotFoundError:
                with pd.ExcelWriter(f"{self.srf}") as writer: # pylint: disable=abstract-class-instcustomiated
                    final_summary.to_excel(writer, sheet_name = "Worst Case SAR", index = False)
            
            else:
                with pd.ExcelWriter(f"{self.srf}", mode = "a", if_sheet_exists = "overlay") as writer: # pylint: disable=abstract-class-instcustomiated
                    final_summary.to_excel(writer, sheet_name = "Worst Case SAR", startrow = writer.sheets["Worst Case SAR"].max_row, index = False, header = False)
            
            return final_summary

        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Not Wi-Fi Summary Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)
    
    def smtx_tech_results(self):
        try:
            smtx_tech_results = [self.data[self.data["Tech"] == self.techlist[tech]] for tech in range(len(self.techlist))]
            
            smarttx_results = [smtx_tech_results[tech].filter(items = ["Tech", "Antenna(s)", "RF Exposure Condition", "Mode(s)", "Power Mode(s)", "Dist. (mm)", "Test Position(s)", "Channel", "Freq. (MHz)", "RB Allocation", "RB Offset", "TuP Limit (dBm)", "Meas. (dBm)", "1-g Meas. (W/kg)", "1-g Scaled (W/kg)", "8-g Meas. (W/kg)", "8-g Scaled (W/kg)", "10-g Meas. (W/kg)", "10-g Scaled (W/kg)", "APD Meas. (W/m2)", "APD Scaled (W/m2)"]) for tech in range(len(smtx_tech_results))]
            
            stx_drop_zero = [smarttx_results[zero][smarttx_results[zero]["1-g Meas. (W/kg)"] != 0] for zero in range(len(smarttx_results))]
            
            smarttx_results = [stx_drop_zero[blank].dropna(subset = ["1-g Meas. (W/kg)"]) for blank in range(len(stx_drop_zero))]
            
            final_stx = pd.concat(smarttx_results)
            
            try:
                abs_filepath_stx = Path(f"{self.stx}").resolve(strict = True)
            
            except FileNotFoundError:
                with pd.ExcelWriter(f"{self.stx}") as writer: # pylint: disable=abstract-class-instantiated
                    final_stx.to_excel(writer, sheet_name = "Worst Case SAR", index = False)
            
            else:
                with pd.ExcelWriter(f"{self.stx}", mode = "a", if_sheet_exists = "overlay") as writer: # pylint: disable=abstract-class-instantiated
                    final_stx.to_excel(writer, sheet_name = "Worst Case SAR", startrow = writer.sheets["Worst Case SAR"].max_row, index = False, header = False)
            
            return final_stx
        
        except Exception as e:
            logger = logging.getLogger("\nSAR GUI: Not Wi-Fi Smart Transmit Module")
            logger.setLevel(logging.ERROR)
            
            lfh = logging.FileHandler(os.path.join(self.log, "error.log"))
            lfh.setLevel(logging.ERROR)
            
            formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            lfh.setFormatter(formatter)
            
            logger.addHandler(lfh)
            
            logger.exception(e)