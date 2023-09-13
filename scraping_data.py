import os
import requests
import time
import random
import numpy as np
import pandas as pd

from urllib.parse import parse_qs, urlparse, unquote
from openpyxl import load_workbook
from collections import defaultdict
from bs4 import BeautifulSoup

with open("modified_urls.txt", "r") as file:
    modified_urls = [line.strip() for line in file]

data_folder = "Data"
if not os.path.exists(data_folder):
    os.makedirs(data_folder)

sheet_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

def get_moon_phase(group):
    if "Full Moon" in group["Data"].values:
        group["Moon Phases"] = 1
    elif "First Qtr" in group["Data"].values or "Last Qtr" in group["Data"].values:
        group["Moon Phases"] = 2
    elif "New Moon" in group["Data"].values:
        group["Moon Phases"] = 3
    else:
        group["Moon Phases"] = 0
    return group

errors2 = []

start_idx = int(os.getenv("START_IDX"))
end_idx = int(os.getenv("END_IDX"))

end_idx = min(end_idx, len(modified_urls))

for url in modified_urls[start_idx:end_idx]:
    params = parse_qs(urlparse(url).query)

    city_state_data = unquote(params["comb_city_info"][0]).split(",")[:2]
    city = city_state_data[0].replace("+", " ")
    state = city_state_data[1].strip().replace("+", " ")

    month = params["month"][0]
    year = params["year"][0]

    month_mapping = {
        "1": "January",
        "2": "February",
        "3": "March",
        "4": "April",
        "5": "May",
        "6": "June",
        "7": "July",
        "8": "August",
        "9": "September",
        "10": "October",
        "11": "November",
        "12": "December"
    }

    month_name = month_mapping.get(month)

    if not os.path.exists(f"{data_folder}/{state}_{year}.xlsx"):
        with pd.ExcelWriter(f"{data_folder}/{state}_{year}.xlsx", engine="openpyxl") as writer:
            for sheet in sheet_names:
                df_empty = pd.DataFrame(columns=["Date", "State", "City", "Moon Phases", "Data", "Time"])
                df_empty.to_excel(writer, sheet_name=sheet, index=False)

    max_retries = 3
    success = False
    attempts = 0

    while not success and attempts < max_retries:
        try:
            response = requests.get(url)
            soup = BeautifulSoup(response.content, "html.parser")
            table = soup.select_one("table[style='width:96%; margin-left:4px; margin-bottom:10px; border-collapse:collapse; border-spacing:0; border:2px solid black; ']")

            data = []
            for row in table.select("tr")[1:]:
                cells = row.select("td")
                for cell in cells:
                    day_data = {}
                    day_number_elements = cell.select("span.daynum")
                    if day_number_elements:
                        day_data["Day"] = day_number_elements[0].text
                        events = cell.get_text('\n').split('\n')[1:]
                        event_dict = defaultdict(list)
                        for event in events:
                            if event and ": " in event:
                                event_parts = event.split(": ")
                                event_dict[event_parts[0]].append(event_parts[1])
                        for k, v in event_dict.items():
                            day_data[k] = ", ".join(v)
                        day_data["City"] = city
                        day_data["State"] = state
                        data.append(day_data)

            df = pd.DataFrame(data)

            for col in df.columns:
                if df[col].dtype == "object" and df[col].str.contains(",").any():
                    splits = df[col].str.split(",", expand=True)
                    
                    for i in range(splits.shape[1]):
                        df[f"{col}_{i+1}"] = splits[i]
                        
                    df.drop(col, axis=1, inplace=True)

            df_melted = df.melt(
                id_vars=["Day", "State", "City"], 
                value_vars=list(df.drop(["Day", "State", "City"], axis=1).columns),
                var_name="Data",
                value_name="Time"
            )
            df_melted = df_melted.replace("none", np.nan)
            df_melted.dropna(inplace=True)
            df_melted = df_melted.groupby("Day").apply(get_moon_phase).reset_index(drop=True)
            df_melted["Day"] = pd.to_datetime(str(month) + "-" + df_melted["Day"].astype(str) + "-" + str(year)).dt.date
            df_melted["Time"] = df_melted["Time"].str.strip().replace("24:00", "00:00")
            df_melted["Time"] = pd.to_datetime(df_melted["Time"], format="%H:%M").dt.time
            df_melted.rename(columns={"Day": "Date"}, inplace=True)
            df_melted = df_melted[["Date", "State", "City", "Moon Phases", "Data", "Time"]]
            df_melted = df_melted.sort_values(["Date", "State", "City", "Time"])
            df_melted["Data"] = df_melted["Data"].str.split("_").str[0]
            df_melted.reset_index(drop=True, inplace=True)

            wb = load_workbook(f"{data_folder}/{state}_{year}.xlsx")
            ws = wb[month_name]

            for index, row in df_melted.iterrows():
                ws.append(row.values.tolist())

            wb.save(f"{data_folder}/{state}_{year}.xlsx")

            print(f"{city}, {state} has been stored in the {month_name} sheet of {data_folder}/{state}_{year}.xlsx\n")

            success = True

        except requests.exceptions.RequestException as e:
            print(f"Error with URL {url}: {e}")
            errors2.append(url)
            attempts += 1

with open("errors2.txt", "a") as file:
    for error in errors2:
        file.write("%s\n" % error)