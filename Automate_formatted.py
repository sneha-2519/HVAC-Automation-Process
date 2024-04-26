import pandas as pd
import re
from tqdm import tqdm
import os
import sys
import openpyxl
from openpyxl import load_workbook
import xlrd
from datetime import datetime
from Automate_JSON import main as creation_json

file_path_to_output = sys.argv[1]
brands_mapping_location = sys.argv[2]
map_api_location = sys.argv[3]
map_location = sys.argv[4]

def main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):
    os.mkdir(file_path_to_output + "/FORMATTED")
    files = [x for x in os.listdir(file_path_to_output + "/RAW")]

    for file in tqdm(files, total=len(files)):
        df = pd.read_excel(file_path_to_output + "/RAW/" + file , sheet_name="Sheet1", keep_default_na=False)
        
        if ('Heat Pumps (Ducted)-Raw.xlsx' in file) and ('Central Air Conditioners (Ducted)-Raw.xlsx' in file):
            #data_dict = df.to_dict()
            df = df.where(pd.notnull(df), None)
        
            na_pattern = re.compile(r"^N/A$|^N\\?A$|^None$", re.I)
            def remove_na(text):
                if text is None or re.search(na_pattern, str(text)):
                    return None
                elif text == '[]':
                    return None
                else:
                    return text
            
            def skip_none(f):
                def wrapped(text):
                    if text:
                        return f(text)
                    else:
                        return None
                return wrapped
            
            def remove(text):
                chars = ['\\n', '\\r', '\\t', u'\u2022', u'\\u0040', u'\\xa0', '\\V']
                quotes = ['“', '”', '″']
                for ch in chars:
                    text = str(text).replace(ch, ' ')
                for q in quotes:
                    text = str(text).replace(q, '"')    
                return text    
                 
            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text
            
            for column in list(df.columns):
                df[column] = df[column].map(remove_na).map(remove).map(trim_space)
            
            
            column_name_map= pd.read_excel(map_location,sheet_name="Sheet1", na_filter=False) 
            column_names={}
            for index, row in column_name_map.iterrows():
                column_names[row['orig_column_name']] = row['rename'] 
            df = df.rename(column_names, axis=1)      
                
            brand_map=pd.read_excel(brands_mapping_location,sheet_name="brands-mapping", na_filter=False) 
            brand_names={}
            for index, row in brand_map.iterrows():
                brand_names[row['brand_name']] = row['brand_rename'] 
            
            df = df.rename(brand_names, axis=1)
            
        
            df['category'] = ''
            df['energy-star-model-name']=''	
            df['energy-star-model-number']=''
        
            #update date
            def add_new(row):
                row['category'] = ''   
                row['category'] = "HVAC/Mechanical"
                row['energy-star-model-name'] = str(row['name'])
                row['energy-star-model-number'] = str(row['name'])
                return row
            
            df = df.apply(add_new, axis=1)
            
            #add timestamp
            df['timestamp'] = ''
            #update date
            def add_timestamp(row):
                row['timestamp'] = ''
                row['timestamp'] = "2024-04-03T00:00:00Z"
                return row
            
            df = df.apply(add_timestamp, axis=1)
        
            
            #add sku
            df['sku'] = ''
            
            def sku_add(row):
                row['sku'] = ''
                if row['name']:
                    name = str(row['name']).replace(" ", "-")
                    error=re.sub(r"[\r\n\t\x07\x0b\xa0\u0040]", " ", name).lower()
                    nama_single = error.replace("'", "")
                    nama_double = nama_single.replace('"','')
                    removeSpecialChars = nama_double.translate ({ord(c): "-" for c in "!@#$）%^（&*()•–＃—”“’‘[]{}';%:,./<>?\|`~-=_+"})
                    remove_space =  re.sub(' +','-',removeSpecialChars)                                  
                    remove_char = re.sub("™|®|", "", remove_space) 
                    mult_hyph = re.sub('-+','-',remove_char) 
                    row['sku'] = mult_hyph.rstrip('-').lstrip('-')          
                return row
            
            df = df.apply(sku_add, axis=1)
            
        
            
            #trim space
            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text
            
            #remove duplicate  
            df = df.sort_values(['brand-name', 'name', 'sku'],ascending=True)
            df.drop_duplicates(subset=['brand-name', 'name'], keep='last', inplace = True)
            
            df['dup_rank'] = (df.groupby(['brand-name', 'sku']).cumcount().add(1)).astype(int)
            
            
            rep_df = df.groupby(['brand-name','sku'])['dup_rank'].agg(max) > 1
            dup_sku = [i for i in rep_df.index if rep_df[i]]
            
            
            def distinct_sku(row):
                sku_bra = row['sku']
                bra =row['brand-name']
                str_brand_sku=str(bra+sku_bra)
                for p in dup_sku:
                    join_list=[''.join(p)]
                    if str_brand_sku in join_list:  
                        row['sku'] = ('-').join([row['sku'],str(row['dup_rank'])]) 
                return row
            
            df = df.apply(distinct_sku, axis=1)
            
            def add_new1(row, file_name):
                if 'sku' not in row:
                    row['sku'] = ''
                if file_name == "Central Air Conditioners (Ducted)":
                    row['subcategory'] = "Air Conditioning, Central"
                    row['sku'] = "cac-" + str(row['sku'])
                elif file_name == "Heat Pumps (Ducted)":
                    row['subcategory'] = "Heat Pumps, Air-Source"
                    row['sku'] = "hpd-" + str(row['sku'])
                return row
            #df = df.apply(lambda row: add_new1(row, file), axis=1) 
            df = df.apply(lambda row: add_new1(row, "Central Air Conditioners (Ducted)" if 'Central Air Conditioners (Ducted)' in file else "Heat Pumps (Ducted)"), axis=1)
            
            df = pd.read_excel(os.path.join(file_path_to_output, file), sheet_name="Sheet1", keep_default_na=False, engine='openpyxl')
            today_date = datetime.today().strftime("%Y-%m-%d")
            df.to_excel(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}-up.xlsx", index=False)
        
        
        else:
            df = df.where(pd.notnull(df), None)

            na_pattern = re.compile(r"^N/A$|^N\\?A$|^None$", re.I)

            def remove_na(text):
                if text is None or re.search(na_pattern, str(text)):
                    return None
                elif text == '[]':
                    return None
                else:
                    return text

            def trim_space(text):
                text = re.sub('<[^<]+?>', '', str(text))
                text = re.sub("\s+", " ", str(text).strip())
                return text

            for column in list(df.columns):
                df[column] = df[column].map(remove_na).map(trim_space)
            column_name_map = pd.read_excel(map_api_location, sheet_name="Sheet1", na_filter=False)
            column_names = {}
            for index, row in column_name_map.iterrows():
                column_names[row['orig_column_name']] = row['rename']
            df = df.rename(column_names, axis=1)

            brand_map = pd.read_excel(brands_mapping_location, sheet_name="brands-mapping", na_filter=False)
            brand_names = {}
            for index, row in brand_map.iterrows():
                brand_names[row['brand_name']] = row['brand_rename']
            df['brand-name'].replace(brand_names, inplace=True)

            def sku_add(row):
                row['sku'] = ''
                if row['energy-star-model-number']:
                    name = str(row['energy-star-model-number']).replace(" ", "-")
                    error = re.sub(r"[\r\n\t\x07\x0b\xa0\u0040]", " ", name).lower()
                    nama_single = error.replace("'", "")
                    nama_double = nama_single.replace('"', '')
                    removeSpecialChars = nama_double.translate({ord(c): "-" for c in
                                                                "!@#$）%^（&*()•–＃—”“’‘[]{}';%:,./<>?\|`~-=_+"})
                    remove_space = re.sub(' +', '-', removeSpecialChars)
                    remove_char = re.sub("™|®|", "", remove_space)
                    mult_hyph = re.sub('-+', '-', remove_char)
                    row['sku'] = mult_hyph.rstrip('-').lstrip('-')
                return row

            df = df.apply(sku_add, axis=1)

            def add_new(row, file_name):
                row['category'] = "HVAC/Mechanical"

                if file_name.lower() == "mini-split-air-conditioners":
                    row['type'] = "Mini-Split AC"
                    row['subcategory'] = "Ductless Heating and Cooling"
                    row['sku'] = "msac-" + row['sku']

                elif file_name.lower() == "room air conditioners":
                    row['subcategory'] = "Air Conditioning, Room"

                elif file_name.lower() == "boilers":
                    row['type'] = "Residential"
                    row['subcategory'] = "Boilers"

                elif file_name.lower() == "commercial boilers":
                    row['subcategory'] = "Boilers"
                    row['type'] = 'Commercial'

                elif file_name.lower() == "geothermal heat pumps":
                    row['subcategory'] = "Heat Pumps, Geothermal or Ground-Source"

                elif file_name.lower() == "ventilating fans":
                    row['subcategory'] = "Ventilation Fans"
                    
                elif file_name.lower() == "furnaces":
                    row['subcategory'] = "Furnaces"

                return row

            df = df.apply(lambda row: add_new(row, file), axis=1)

            df['url'] = df.apply(lambda row: f"https://www.energystar.gov/productfinder/product/certified-{file.lower().replace(' ', '-')}/details/{row['energy-star-id']}", axis=1)
            
            
            def add_bool(df):
                for i, row in df.iterrows():
                    if "energy-star-lamp-included" in row:
                        if row['energy-star-lamp-included']=="Yes":
                            row['energy-star-lamp-included']=True
                        if row['energy-star-lamp-included']=="No":
                            row['energy-star-lamp-included']=""            
                    if "meets-most-efficient-criteria-2024-2024" in row:
                        if row['meets-most-efficient-criteria-2024']=="Yes":
                            row['meets-most-efficient-criteria-2024']=True
                        if row['meets-most-efficient-criteria-2024']=="No":
                            row['meets-most-efficient-criteria-2024']=""    
        
                    if "low-noise" in row:
                        if row['low-noise']=="Yes":
                            row['low-noise']=True
                        if row['low-noise']=="No":
                            row['low-noise']=""   
                            
                    if "variable-speed-compressor" in row:
                        if row['variable-speed-compressor']=="Yes":
                            row['variable-speed-compressor']=True
                        if row['variable-speed-compressor']=="No":
                            row['variable-speed-compressor']=""                   
        
                    if "date-available-on-market" in row:
                        date_ava = row['date-available-on-market']
                        row['date-available-on-market']=date_ava.replace("T00:00:00.000", "T00:00:00Z")
        
                    if "starts" in row:
                        date_ava = row['starts']
                        row['starts']=date_ava.replace("T00:00:00.000", "T00:00:00Z")          
                return df
            df = add_bool(df)


            rem_brand2=['TeK','Ace','Toshiba','Philips','PHILIPS','/', 'SAVIN', 'RICOH', 'LANIER', 'HPE', 'Acer','Compumax','dynabook', '®','-','?','Dell EMC','EMC','HP']
            def name_add(row):
                row['name'] = ''
                if 'energy-star-model-name' in row and row['energy-star-model-name']:
                    name = str(row['energy-star-model-name'])
                    row['name'] = name
                return row
            df = df.apply(name_add, axis=1)

            def remv_brand2(df):
                if 'name' in df.columns:
                    for i, row in df.iterrows():
                        if row['brand-name']:
                            for b in rem_brand2:
                                if b in row['name']:
                                    row['name'] = row['name'].lstrip(b).strip()   
                                if b in row['energy-star-model-number']:
                                    row['energy-star-model-number'] = row['energy-star-model-number'].lstrip(b).strip()                      
                return df
            
            df = remv_brand2(df)
            
            df['timestamp'] = "2024-03-04T00:00:00Z"

            df = df.reindex(columns=(['timestamp', 'brand-name', 'sku', 'name', 'energy-star-model-name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'energy-star-id', 'additional-model-information', 'energy-star-partner', 'markets'] +
                                      [a for a in df.columns if a not in ['timestamp', 'brand-name', 'sku', 'energy-star-model-name', 'name', 'energy-star-model-number', 'category', 'subcategory', 'type', 'energy-star-id', 'additional-model-information', 'energy-star-partner', 'markets']]))

            df = df.sort_values(['brand-name', 'name', 'sku'], ascending=True)
            df.drop_duplicates(subset=['brand-name', 'energy-star-model-name', 'sku', 'energy-star-id'], keep='last', inplace=True)

            df['dup_rank'] = df.groupby(['brand-name', 'sku']).cumcount().add(1).astype(int)
            rep_df = df.groupby(['brand-name', 'sku'])['dup_rank'].agg(max) > 1
            dup_sku = [i for i in rep_df.index if rep_df[i]]

            def distinct_sku(row):
                sku_bra = row['sku']
                bra = row['brand-name']
                str_brand_sku = str(bra + sku_bra)
                if str_brand_sku in dup_sku:
                    row['sku'] = '-'.join([row['sku'], str(row['dup_rank'])])
                return row

            df = df.apply(distinct_sku, axis=1)
            today_date = datetime.today().strftime("%Y-%m-%d")
            df.to_excel(file_path_to_output + "/FORMATTED/" + f"{file.replace('-Raw.xlsx', '')}-{today_date}-up.xlsx", index=False)

            #df.to_excel(file_path_to_output + "/FORMATTED/" + file.replace("-RAW.xlsx", "") + "-up.xlsx", index=False)
            if file.endswith("-up.xlsx"):
                file_path = os.path.join(file_path_to_output + "/FORMATTED/", file)
                print("file path",file_path)
                print("file path",file_path)
                wb = openpyxl.load_workbook(file_path) # Open the Excel 
                ws = wb.worksheets[0]
                worksheet = wb.active # Get the active worksheet
                worksheet.freeze_panes = "A2"  # Freezes everything above row 2 (including row 1)
                worksheet.auto_filter.ref = worksheet.dimensions  # Applies filter to the entire
                # Left-align the text in the first row (assuming you want this for all columns)
                for cell in worksheet[1]:
                    alignment = openpyxl.styles.Alignment(horizontal="left")
                wb.save(os.path.join(file_path_to_output + "/FORMATTED/", file)) 
            
    print("\n FORMATTED UP FILES DONE\n")    
    creation_json(file_path_to_output, brands_mapping_location, map_api_location, map_location,last_month_file_location)
#if __name__ == "__main__":
#    main()
