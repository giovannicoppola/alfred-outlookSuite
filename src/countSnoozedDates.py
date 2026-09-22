#!/usr/bin/env python
# encoding: utf-8
# Wednesday, May 5, 2021
# added number of days on Thursday, February 17, 2022, 2:35 PM
# updated and refreshed on Friday, June 2, 2023

import os
import sys
from consts import OUTLOOK_SNOOZER_FILE
from datetime import date, datetime, timedelta
import json

if not os.path.exists(OUTLOOK_SNOOZER_FILE):
    print("No snoozed emails.")
    sys.exit(0)

with open(OUTLOOK_SNOOZER_FILE, "r") as f:
    myJSON = json.load(f)

if not myJSON:
    print("No snoozed emails.")
    sys.exit(0)

# Get the current date
today = date.today()

# Create an empty report string
report = ""

# Sort the dictionary based on the date values
sorted_dict = sorted(myJSON.items(), key=lambda x: x[1])

# Store, per unique date, its count and its (already-correct) days-from-today.
date_counts = {}

# Iterate over the sorted key-value pairs
for key, value in sorted_dict:
    # Convert the date string to a datetime object
    date_obj = datetime.strptime(value, "%Y-%m-%d").date()

    # Calculate the number of days from today
    days_difference = (date_obj - today).days

    # Format the date as "Weekday, Month Day"
    formatted_date = date_obj.strftime("%A, %B %d")

    # Add or increment the count; keep the days value from the original date so
    # we don't re-parse it against the current year (which broke across year
    # boundaries and crashed on Feb 29 in a non-leap year).
    entry = date_counts.setdefault(formatted_date, {"count": 0, "days": days_difference})
    entry["count"] += 1

# Construct the report string with one line per date
for formatted_date, entry in date_counts.items():
    report += f"{formatted_date} ({entry['days']}): {entry['count']}\n"

# Print the report string
print(report)
	
	
