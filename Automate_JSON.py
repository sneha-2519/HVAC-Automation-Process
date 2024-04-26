import pandas as pd
import json
import os
from datetime import datetime
import xlrd 
import sys
from tqdm import tqdm
from Comparing_excel_up import main as formatting_upExcel
#file_path_to_output = sys.argv[1]
#brands_mapping_location = sys.argv[2]
#map_api_location = sys.argv[3]
#map_location = sys.argv[4]

def main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):
    print("\n  Creation of JSON files: ")
    os.mkdir(file_path_to_output + "/JSON")

    files = [x for x in os.listdir(file_path_to_output + "/FORMATTED")]
    #float_columns = ['seer2', 'eer2', 'thermal-efficiency-te']

    # Iterate over each file in the folder
    for file_name in tqdm(files, total=len(files)):
        df = pd.read_excel(os.path.join(file_path_to_output, "FORMATTED", file_name), sheet_name="Sheet1", keep_default_na=False)

        if file_name.endswith('-up.xlsx'):  # Check if the file is an Excel file
            #df = pd.read_excel(file_path_to_output + "/FORMATTED/" + file_name , sheet_name="Sheet1", keep_default_na=False)

            
            data_dict = df.to_dict()
            df = df.where(pd.notnull(df), None)
            
            data = df.to_dict()
            data_list = []
            data_dict = {}
            
            # Iterate through rows in the DataFrame
            for i in range(len(df)):
                data_dict = {}
                data_dict['disabled'] = False 
                data_dict['energy-star'] = True  
                if data['timestamp'][i]:
                    data_dict['timestamp'] = str(data['timestamp'][i]) 
                if data['brand-name'][i]:
                    data_dict['brand-name'] = str(data['brand-name'][i])         
                if data['sku'][i]:
                    data_dict['sku'] = str(data['sku'][i])  
                if data['name'][i]:
                    data_dict['name'] = str(data['name'][i])         
                if data['energy-star-model-name'][i]:
                    data_dict['energy-star-model-name'] = str(data['energy-star-model-name'][i])   
                if data['energy-star-model-number'][i]:
                    data_dict['energy-star-model-number'] = str(data['energy-star-model-number'][i])        
                if 'category' in data.keys():
                    if data['category'][i]:    
                        data_dict['category'] = [data['category'][i]] 
                if 'subcategory' in data.keys():
                    if data['subcategory'][i]:    
                        data_dict['subcategory'] = [data['subcategory'][i]] 
                if 'type' in data.keys():
                    if data['type'][i]:    
                        data_dict['type'] = str(data['type'][i])
                if data['energy-star-id'][i]:
                    data_dict['energy-star-id'] = str(data['energy-star-id'][i])            
                
                
                
                data_dict["energy-star-certificate"] = []
                cert ={}
                if data['energy-star-id'][i]:
                    cert['id'] = str(data['energy-star-id'][i]) 
                if 'starts' in data.keys():
                    if data['starts'][i].strip():
                        cert['starts'] = str(data['starts'][i])
                if data['url'][i]:
                    cert['url'] = str(data['url'][i])         
                data_dict["energy-star-certificate"].append(cert)  
                
                x_str=['additional-ct-device-model-numbers', 'additional-model-information', 'ahri-reference-number', 'alternate-energy-star-lamps-esuid', 'boiler-application', 'boiler-control-type', 'broadband-connection-needed-for-demand-response', 'can-integrate-hot-water-heating', 'capable-of-two-way-communication', 'casement-window', 'cold-climate', 'communication-hardware-architecture', 'communication-method-other', 'communication-standard-application-layer', 'compressor-staging', 'connected-capability', 'connected-capable', 'connected-functionality', 'connect-using', 'cooling-capacity', 'cooling-capacity-range', 'cop-at-5f', 'correlated-color-temperature-kelvin', 'ct-device-brand-name', 'ct-device-brand-owner', 'ct-device-communication-method', 'ct-device-model-name', 'ct-device-model-number', 'ct-product-heating-and-cooling-control-features', 'demand-response-product-variations', 'demand-response-summary', 'depth-inches', 'direct-on-premises-open-standard-based-interconnection', 'dr-protocol', 'duct-size', 'eer-range-btu-wh', 'efficiency-afue', 'energy-star-lamp-partner', 'energy-star-model-identifier', 'energy-star-partner', 'family-id', 'fan-lamp-model-number', 'features', 'fuel-type', 'furnace-is-energy-star-certified-in', 'heating-capacity', 'heating-capacity-at-17-f-btu-h', 'heating-capacity-at-47-f', 'heating-capacity-at-47-f-btu-h', 'heating-capacity-at-5-f', 'heating-capacity-at-5-f-btu-h', 'heating-mode', 'height-inches', 'hspf', 'hspf-range-btu-wh', 'indoor-unit-model-number', 'installation-capabilities', 'installation-mounting-type', 'lighting', 'lighting-technology-used', 'meets-peak-cooling-requirements', 'network-security-standards', 'notes', 'number-of-speeds', 'other-heating-and-cooling-control-features', 'outdoor-unit-brand-name', 'primary-communication-module-device-brand-name-and-model-number', 'product-class', 'refrigerant-type', 'refrigerant-type-gwp', 'reverse-cycle', 'seer-range-btu-wh', 'sound-level-sones', 'special-features-dimming-motion-sensing-etc', 'support-bracket', 'tax-credit-eligible', 'tax-credit-eligible-cac-national', 'tax-credit-eligible-heat-pumps-north', 'tax-credit-eligible-heat-pumps-south', 'voltage-volts', 'weight-lbs', 'width-inches']
                
                for j in x_str:
                    if j in data.keys():
                        if data[j][i]:
                            data_dict[j]  = str(data[j][i]) 
                
                x_int=['aeu', 'airflow-1-cfm', 'airflow-2-cfm', 'airflow-3-cfm', 'bathroom-utility-room-airflow-at-025-in-wg', 'boiler-full-load-input-rate', 'boiler-turndown-ratio', 'color-rendering-index-cri', 'combined-energy-efficiency-ratio-ceer', 'cool-cap', 'cooling-capacity-kbtu-h', 'cop-rating', 'cop-rating-at-17-degrees', 'cop-rating-at-47-degrees', 'eer2-rating-btu-wh', 'eer-rating-btu-wh', 'efficacy-1-cfm-w', 'efficacy-2-cfm-w', 'efficacy-3-cfm-w', 'energy-efficiency-ratio-eer', 'energy-star-lamp-esuid', 'ieer-rating', 'le-measured', 'light-out', 'light-source-life-hours', 'merv-of-in-line-fan-filter', 'network-standby-average-power-consumption', 'percent-less-energy-use-than-us-fed-standard', 'power-factor', 'seer2-rating-btu-wh', 'seer-rating-btu-wh', 'static-temperature-accuracy', 'thermal-efficiency-te']
                
                for k in x_int:
                    if k in data.keys():
                        if data[k][i]:
                            data_dict[k]  = int(float(data[k][i]))
                            
                if 'markets' in data.keys():
                    if data['markets'][i]:
                        if ',' in data['markets'][i]:
                            mark = data['markets'][i].split(",")
                            data_dict['markets'] = [x.strip(' ') for x in mark]  
                        else:
                            data_dict['markets'] = str(data['markets'][i])
                            
                if 'upc' in data.keys():
                    if data['upc'][i]:
                        if ';' in data['upc'][i]:
                            upc_str = str(data['upc'][i]).split(';')
                            upc_trim = [x.strip(' ') for x in upc_str]       
                            data_dict['upc'] = upc_trim
                        else:
                            data_dict['upc'] = str(data['upc'][i])
                            
                if 'date-available-on-market' in data.keys():
                    if data['date-available-on-market'][i]:    
                        data_dict['date-available-on-market'] = str(data['date-available-on-market'][i])

                if 'meets-most-efficient-criteria-2024' in data.keys():
                    if data['meets-most-efficient-criteria-2024'][i]:
                        data_dict['meets-most-efficient-criteria-2024'] = bool(data['meets-most-efficient-criteria-2024'][i])
                        
                if 'low-noise' in data.keys():
                    if data['low-noise'][i]:
                        data_dict['low-noise'] = bool(data['low-noise'][i])
                        
                if 'variable-speed-compressor' in data.keys():
                    if data['variable-speed-compressor'][i]:
                        data_dict['variable-speed-compressor'] = bool(data['variable-speed-compressor'][i])
                        
                if 'energy-star-lamp-included' in data.keys():
                    if data['energy-star-lamp-included'][i]:
                        data_dict['energy-star-lamp-included'] = bool(data['energy-star-lamp-included'][i])
                
                sort_dict=dict(sorted(data_dict.items()))
                data_list.append(sort_dict)

            json_filename = os.path.join(file_path_to_output + "/JSON/" + file_name.replace('-up.xlsx', '.json'))
            with open(json_filename, 'w', encoding='utf-8') as fout:
                json.dump(data_list, fout, indent=1, ensure_ascii=False, default=str) 
                
    print("\n Creation of JSON Files are completed. ")
        
    formatting_upExcel(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)