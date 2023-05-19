import requests
import pandas as pd
from tqdm import tqdm

# a = your excel file's directory
a = "sample.xlsx"
report = pd.read_excel(a)
report = report[["recordID","datetimeISO8601", "date","lat","lon"]]

#NOTE THAT ALL DATE FIELDS SHOULD BE GMT+0
#If you receive timeout error, reduce your data to about 700 rows.
#recordID field is for you to match the results with your database.
#datetimeISO8601 format is %Y-%m-%dT%H:%M (for example: "2010-05-22T08:00")
#date field is the date of datetimeISO8601 column: 2010-05-22
#you can easily convert your date and time columns into ISO8601 format via codes below:
"""df = your excel
df["datetimeISO8601"] = df['your date column']  + pd.to_timedelta(df['your time column'].astype(str))
df["datetimeISO8601"] = df["datetimeISO8601"].dt.floor("H").dt.strftime("%Y-%m-%dT%H:%M")"""
#lat is latitude. for example: "40,1532248846051"
#lon is longitude. for example: "26,407608983852697"


baseURL="https://archive-api.open-meteo.com/v1/archive?"

print("the loop has been started")
for row in tqdm(report.index, total=len(report.index)):
    try:
        lat = str(report.loc[row, "lat"])
        lon = str(report.loc[row, "lon"])
        date = str(report.loc[row, "date"].date())
        datetimeISO8601 = str(report.loc[row, "datetimeISO8601"])
        templateURL = baseURL + f"latitude={lat}&longitude={lon}&start_date={date}&end_date={date}&hourly=temperature_2m"
        #if you want to get different data, for example: for apparent temperature instead of real temperature, put "apparent_temperature" instead of "temperature_2m".

        response = requests.get(templateURL).json()
        data = pd.DataFrame(response["hourly"])
        data["time"] = data["time"].astype(str)
        srow = data.loc[data["time"] == datetimeISO8601]
        temp = srow["temperature_2m"].values[0]
        report.at[row, "temperature"] = temp
    except IndexError:
        report.at[row, "temperature"] = "NaN"
        continue
print("saving...")
report.to_excel("result.xlsx", index=False)
print("The process has been completed.")