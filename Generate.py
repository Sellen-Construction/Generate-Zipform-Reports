""" 
    Generate Annual ZIP Form CSV Reports
    Version 1.2

    Created by: Michael Rice

    Changelog:
        1.0 (10/20/2020): Initial release. Generate reports for the year specified in get_year and save the results to "{get_year} report.csv"
        1.1 (10/20/2020): Prompt for get_year at runtime. Removes the need to modify the code to change the year
        1.2 (10/21/2020): Include "Workstations" and "Other" in the tally. These are legacy values included to capture historic data
"""

# Required imports. Ensure requests and configparser have been installed before running
#python3 -m pip install -r requirements.txt
import requests, json, configparser, csv, calendar, datetime

# Generate a report for the following year
current_year = datetime.datetime.now().year
input_year = input(f"Generate a report for which year? (Default: {current_year}): ")

# If an invalid year was given, keep prompting for a valid year
while str.isdigit(input_year) and len(input_year) != 4:
    print("Not a valid year. Please enter a valid 4 digit year")
    input_year = input(f"Generate a report for which year? (Default: {current_year}): ")

# 
get_year = input_year if input_year else current_year

# Set various values. These were designed for version 1.0 of the Asana API
root_URL = "https://app.asana.com/api/1.0"
opt_fields = "name,created_at,custom_fields.enum_value.name"
project_id = "1146674956832113"
device_gid = "1152024058544808"
accessory_gid = "1152024058544805"
workstation_gid = "1152074579726438"
request_type_gid = "1152024058544802"

# Load the encrypted API key from the config file
config = configparser.ConfigParser()
config.read('config.txt')
key = config['Default']['key']
headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Authorization': 'Bearer ' + key
    }

# Load all of the device types that exist in Asana. 
#If those types change, this report will automatically grab the most current list
device_types = []
types_source = json.loads(requests.get(f"{root_URL}/custom_fields/{device_gid}", headers = headers).text)
for device in types_source['data']['enum_options']:
    if device['enabled']:
        device_types.append(device['name'])

# Prepare count totals for each month
month_counts = {}
for month in range(1, 13):
    month = str(month)
    if len(month) < 2:
        month = "0" + month
    month_counts[month] = 0

# Prepare accessory, workstation, and device counts for each month
accessory_counts = month_counts.copy()
workstation_counts = month_counts.copy()
other_counts = month_counts.copy()
device_counts = {}
for device in device_types:
    device_counts[device] = month_counts.copy()

# Call the Asana API
report_source = json.loads(requests.get(f"{root_URL}/projects/{project_id}/tasks?opt_fields={opt_fields}", headers = headers).text)

# Now that we have the data from Asana, make sure there were no errors and parse the data
for row in report_source['data']:
    if row['created_at'][0:4] == str(get_year):
        for field in row['custom_fields']:
            # Ensure we have all of the data we need in order to parse
            if "enum_value" in field:
                month = row['created_at'][5:7]
                
                if field['enum_value'] is not None and "name" in field['enum_value']:
                    # Tally all of the requests for each month
                    if field['enum_value']['name'] in device_types and field['gid'] == device_gid:
                        device_counts[field['enum_value']['name']][month] += 1
                    elif field['enum_value']['gid'] == accessory_gid:
                        accessory_counts[month] += 1
                    elif field['enum_value']['gid'] == workstation_gid:
                        workstation_counts[month] += 1
                    else:
                        continue
                    month_counts[month] += 1
                elif field['gid'] == request_type_gid and field['enum_value'] is None:
                    other_counts[month] += 1
                    month_counts[month] += 1
                    continue

# Output the results to a CSV file
with open(f"{get_year} report.csv", "w", newline = '\n') as file:
    output = csv.writer(file)
    
    # Prepare the header row
    row = [get_year]
    for key in month_counts.keys():
        row.append(calendar.month_name[int(key)])
    row.append("Total")
    output.writerow(row)

    # Add the device counts to the file for each month and the totals
    for count in device_counts:
        total = 0
        row = [count]
        for key in device_counts[count].keys():
            row.append(device_counts[count][key])
            total += device_counts[count][key]
        row.append(total)
        output.writerow(row)
    
    # Add the accessory counts to the file for each month and the total
    total = 0
    row = ["Accessories"]
    for month in accessory_counts.keys():
        row.append(accessory_counts[month])
        total += accessory_counts[month]
    row.append(total)
    output.writerow(row)

    # Add the workstation counts to the file for each month and the total
    total = 0
    row = ["Workstation"]
    for month in workstation_counts.keys():
        row.append(workstation_counts[month])
        total += workstation_counts[month]
    row.append(total)
    output.writerow(row)

    # Add the other counts to the file for each month and the total
    total = 0
    row = ["Other"]
    for month in other_counts.keys():
        row.append(other_counts[month])
        total += other_counts[month]
    row.append(total)
    output.writerow(row)

    # Add the totals for each month, as well as the grand total
    total = 0
    row = ["Total"]
    for month in month_counts.keys():
        row.append(month_counts[month])
        total += month_counts[month]
    row.append(total)
    output.writerow(row)