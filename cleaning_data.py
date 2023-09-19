import os
import pandas as pd

target_names = [
    "New Jersey_2022",
    "New Hampshire_2022",
    "Nevada_2023",
    "New Hampshire_2023",
    "Puerto Rico_2023",
    "Virginia_2023",
    "Ohio_2023",
    "New York_2023",
    "Washington county_2023",
    "West Virginia_2023",
    "Oregon_2023",
    "Wyoming_2022",
    "Nevada_2022",
    "North Carolina_2023",
    "New York_2022",
    "Virginia_2022",
    "Rhode Island_2022",
    "U.S. Virgin Islands_2022",
    "U.S. Virgin Islands_2023",
    "Texas_2022",
    "Wisconsin_2023",
    "Ohio_2022",
    "Pennsylvania_2023",
    "Wisconsin_2022",
    "Pennsylvania_2022",
    "Oregon_2022",
    "South Carolina_2022",
    "Washington_2022",
    "Vermont_2023",
    "South Dakota_2022",
    "Nebraska_2023",
    "North Dakota_2023",
    "South Carolina_2023",
    "Utah_2023",
    "Washington_2023",
    "Wyoming_2023",
    "Vermont_2022",
    "Nebraska_2022",
    "Tennessee_2023",
    "New Jersey_2023",
    "Oklahoma_2022",
    "Washington county_2022",
    "Puerto Rico_2022",
    "North Dakota_2022",
    "Utah_2022",
    "New Mexico_2023",
    "West Virginia_2022",
    "Oklahoma_2023",
    "Texas_2023",
    "Rhode Island_2023",
    "New Mexico_2022",
    "North Carolina_2022"
]

# Define folders and target names
folders = ["Data_V2", "Retry_Data_V2"]
final_data_folder = "Final Data"

if not os.path.exists(final_data_folder):
    os.makedirs(final_data_folder)

# For each target name (e.g., Nebraska_2022), search, merge, and save
for target_name in target_names:
    matching_files = []

    for folder in folders:
        for filename in os.listdir(folder):
            if target_name in filename:  # Check if the target name (like Nebraska_2022) is in the filename
                matching_files.append(os.path.join(folder, filename))

    # Dictionary to store data from different sheets
    combined_data = {}

    for filepath in matching_files:
        xls = pd.ExcelFile(filepath)
        for sheet_name in xls.sheet_names:
            if sheet_name not in combined_data:
                combined_data[sheet_name] = []
            combined_data[sheet_name].append(pd.read_excel(xls, sheet_name))

    # For each sheet, concatenate, remove duplicates, and sort
    for sheet_name, dataframes in combined_data.items():
        df = pd.concat(dataframes, ignore_index=True)
        df.drop_duplicates(subset=["Date", "State", "City", "Moon Phases", "Data", "Time"], inplace=True)
        df["Date"] = pd.to_datetime(df["Date"])
        df.sort_values(by=["Date", "State", "City", "Time"], inplace=True)
        combined_data[sheet_name] = df

    # Save the combined data to the "Final Data" folder under the State_Year.xlsx format
    output_filename = os.path.join(final_data_folder, f"{target_name}.xlsx")

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        for sheet_name, df in combined_data.items():
            if df.empty:
                # Create an empty DataFrame with the desired columns
                df = pd.DataFrame(columns=["Date", "State", "City", "Moon Phases", "Data", "Time"])
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Finished merging for {target_name}.xlsx")

print("Merge operation completed")