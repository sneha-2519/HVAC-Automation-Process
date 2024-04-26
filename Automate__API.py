from lxml import html
import requests
import pandas as pd
from sodapy import Socrata
from tqdm import tqdm
import logging
from bs4 import BeautifulSoup
import sys
import os
import re 
from Automate_formatted import main as formatting_main

output_file_location = sys.argv[1]
brands_mapping_location = sys.argv[2]
map_api_location = sys.argv[3]
map_location = sys.argv[4]
last_month_file_location = sys.argv[5]

os.mkdir(output_file_location + "/RAW" )

page = requests.get('https://www.energystar.gov/productfinder/advanced')
tree = html.fromstring(page.content)

# Set the logging level to suppress the warning
logging.getLogger().setLevel(logging.ERROR)

span_elements = tree.xpath("//div//h3[text()='Heating & Cooling']/../ul//" +
                               "li//div//span")

# Initialize finder_url outside the loop
finder_url = None

for span in tqdm(span_elements, total=len(span_elements)):
    file = span.text.strip()
    
    if ('Heat Pumps (Ducted)' in file) and ('Central Air Conditioners (Ducted)' in file):
        api_heat_pumps = "https://www.energystar.gov/productfinder/product/certified-central-heat-pumps/results?page_number="
        api_air_conditioners = "https://www.energystar.gov/productfinder/product/certified-central-air-conditioners/results?page_number="

        def scrape_and_save(api, file_name):
            heat_list = []

            for i in range(29):
                res = api + str(i)
                response = requests.get(res)
                text = response.text
               
                soup = BeautifulSoup(text, 'html.parser')
                rowheading = soup.find_all('div', attrs={'class': 'row'})

                for item in rowheading:
                    heat_dict = {}
                    title = item.find('div', class_='title')
                    if title:
                        name = ' '.join(title.text.strip().split())
                        cleaned_text = ' '.join(name.split())
                        match = re.match(r'^(.*?)\s*-\s*(.*)$', cleaned_text)

                        if match:
                            brand = match.group(1).strip()
                            model = match.group(2).strip()
                        else:
                            print("Unable to extract brand and model.")

                        heat_dict['brand-name'] = brand
                        heat_dict['name'] = model

                    head = ""
                    field = item.find_all('div', attrs={'class': 'field'})

                    for r in field:
                        label = r.find_all('div', attrs={'class': 'label'})

                        for lab in label:
                            head = lab.text.replace("\n", "").strip()

                        value = r.find_all('div', attrs={'class': 'value'})

                        for val in value:
                            value_txt = val.text.replace("\n", "").strip()
                            value_txt = re.sub(r'\s+', ' ', value_txt)
                            heat_dict[head] = value_txt
             
                    if heat_dict:
                        heat_list.append(heat_dict)

            df = pd.DataFrame.from_dict(heat_list)
            df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
            df.drop_duplicates(inplace=True)
            df.to_excel(f"{output_file_location}/RAW/{file}-Raw.xlsx")
            return df
        
    else:
        for a_tag in span.findall("a"):
            if "Finder" in a_tag.text:
                finder_url = "https://www.energystar.gov/productfinder" + \
                    a_tag.get('href')[1:]
            if "API" in a_tag.text:
                gotten_url = a_tag.get('href')
                url = gotten_url.replace("https://dev.socrata.com/foundry/" + "data.energystar.gov/", "")

                client = Socrata("data.energystar.gov", None)
                results = client.get(url, limit=10000)

                df = pd.DataFrame.from_records(results)
                # Removing markets that don't contain the U.S
                df = df[df['markets'].str.contains('United States', na=False)]
                #df.to_excel(
                #    f"{output_file_location}/RAW/{file}-Raw.xlsx")
                
        req_html = requests.get(finder_url)
        if req_html.status_code == 200:
            soup_html = BeautifulSoup(req_html.content, features='lxml')
            site_count = int(soup_html.find_all('div', class_="records-found-small")[0].get_text().strip().replace("\xa0Records Found", ""))
        #if df.shape[0] != site_count:
            print("\n", df.shape[0], site_count, file) 
        df.to_excel(f"{output_file_location}/RAW/{file}-Raw.xlsx")


print("\nRAW FILES DONE\n")
formatting_main(output_file_location, brands_mapping_location, map_api_location, map_location, last_month_file_location)
