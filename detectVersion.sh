#!/bin/bash

# Get the currently logged-in user
loggedInUser=$(scutil  <<< "show State:/Users/ConsoleUser" | awk '/Name :/ && ! /loginwindow/ { print $3 }')

# Check the preference file for the active Outlook version
newOutlookEnabled=$(defaults read /Users/$loggedInUser/Library/Containers/com.microsoft.Outlook/Data/Library/Preferences/com.microsoft.Outlook IsRunningNewOutlook)

if [ "$newOutlookEnabled" == "1" ]; then
    echo "The user is currently using the New Outlook."
else
    echo "The user is currently using the Legacy Outlook."
fi
