import os

def detect_chrome_profile():
    base_path = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    if not os.path.exists(base_path):
        print("Chrome user data folder not found.")
        return None

    profiles = []
    for folder in os.listdir(base_path):
        if os.path.isdir(os.path.join(base_path, folder)):
            if folder == "Default" or folder.startswith("Profile "):
                profiles.append(folder)

    if not profiles:
        print("No Chrome profiles found.")
        return None

    print("Detected Chrome profiles:")
    for i, profile in enumerate(profiles):
        print(f"{i + 1}. {profile}")

    selected_index = int(input("Select a profile to use (number): ")) - 1
    selected_profile = profiles[selected_index]
    
    full_path = os.path.join(base_path)
    print("\n✅ Use these in your script:")
    print(f'options.add_argument("user-data-dir={full_path.replace(os.sep, "/")}")')
    print(f'options.add_argument("profile-directory={selected_profile}")')

detect_chrome_profile()
