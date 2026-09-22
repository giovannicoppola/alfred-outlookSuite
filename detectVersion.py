import os
import subprocess


def get_logged_in_user():
    """Get the currently logged-in user."""
    try:
        # First attempt: Use scutil to get the ConsoleUser
        result = subprocess.run(
            ["scutil", "--get", "ConsoleUser"],
            capture_output=True,
            text=True,
            check=True,
        )
        logged_in_user = result.stdout.strip()
        if logged_in_user and logged_in_user != "loginwindow":
            return logged_in_user
    except subprocess.CalledProcessError:
        # If scutil fails, try an alternative method
        pass

    # Second attempt: Use the environment variable for the user
    logged_in_user = os.getlogin()
    if logged_in_user:
        return logged_in_user

    # If all else fails, return None
    return None


def check_outlook_version(logged_in_user):
    """Check if the user is using the new or legacy version of Outlook."""
    try:
        # Path to the preference file for Outlook
        plist_path = f"/Users/{logged_in_user}/Library/Containers/com.microsoft.Outlook/Data/Library/Preferences/com.microsoft.Outlook"

        # Use the defaults command to read the 'IsRunningNewOutlook' key
        result = subprocess.run(
            ["defaults", "read", plist_path, "IsRunningNewOutlook"],
            capture_output=True,
            text=True,
            check=True,
        )
        # Check if the key value is '1' (New Outlook) or '0' (Legacy Outlook)
        key_value = result.stdout.strip()
        if key_value == "1":
            print("The user is currently using the New Outlook.")
        elif key_value == "0":
            print("The user is currently using the Legacy Outlook.")
        else:
            print(f"Unexpected value for IsRunningNewOutlook: {key_value}")
    except subprocess.CalledProcessError as e:
        print(f"Error reading Outlook preference: {e}")
        print(
            "Unable to determine the active Outlook version. The preference file or key may not exist."
        )


if __name__ == "__main__":
    user = get_logged_in_user()
    if user:
        print(f"Logged-in user: {user}")
        check_outlook_version(user)
    else:
        print("Could not determine the logged-in user.")
