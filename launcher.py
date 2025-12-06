import os
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox
import requests
from packaging import version
import shutil

# Configuration
REPO_OWNER = "fabian-plaehn"
REPO_NAME = "Lohneingabe"
APP_EXECUTABLE_NAME = "Lohneingabe_App.exe"  # The name of the downloaded main app
VERSION_FILE = "version.txt"

def get_current_version():
    """Reads the current version from version.txt."""
    if not os.path.exists(VERSION_FILE):
        return "0.0.0"
    with open(VERSION_FILE, "r") as f:
        return f.read().strip()

def get_latest_release():
    """Fetches the latest release info from GitHub."""
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"Error fetching release info: {e}")
        return None

def download_asset(download_url, target_path):
    """Downloads the file from the given URL."""
    try:
        with requests.get(download_url, stream=True) as r:
            r.raise_for_status()
            with open(target_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        return True
    except Exception as e:
        print(f"Error downloading asset: {e}")
        return False

def launch_app():
    """Starts the main application."""
    if os.path.exists(APP_EXECUTABLE_NAME):
        subprocess.Popen([APP_EXECUTABLE_NAME])
    else:
        # Fallback for development (running from source if exe not found)
        # This part might need adjustment depending on dev setup, but for prod exe it expects the exe.
        # If we are in dev, maybe we try running main.py?
        if os.path.exists("main.py"):
             subprocess.Popen([sys.executable, "main.py"])
        else:
             messagebox.showerror("Error", f"Could not find {APP_EXECUTABLE_NAME}")

def main():
    # Hide the console window if possible (for pyinstaller/nuitka --noconsole)
    # But we might want it for debugging for now.
    
    current_ver = get_current_version()
    print(f"Current version: {current_ver}")
    
    release_info = get_latest_release()
    
    if release_info:
        latest_tag = release_info.get("tag_name", "0.0.0").lstrip("v").split("-")[0]
        print(f"Latest version: {latest_tag}")
        
        if version.parse(latest_tag) > version.parse(current_ver):
            root = tk.Tk()
            root.withdraw() # Hide the main window
            
            if messagebox.askyesno("Update Available", f"A new version ({latest_tag}) is available. Do you want to download and install it?"):
                # Find the asset
                assets = release_info.get("assets", [])
                asset_url = None
                for asset in assets:
                    if asset["name"] == APP_EXECUTABLE_NAME:
                        asset_url = asset["browser_download_url"]
                        break
                
                if asset_url:
                    # Download to a temp file first
                    temp_path = APP_EXECUTABLE_NAME + ".new"
                    if download_asset(asset_url, temp_path):
                        # Close any running instances if necessary? 
                        # Assuming Launcher is running, App is not.
                        
                        # Replace
                        if os.path.exists(APP_EXECUTABLE_NAME):
                             os.remove(APP_EXECUTABLE_NAME)
                        os.rename(temp_path, APP_EXECUTABLE_NAME)
                        
                        # Update version file
                        with open(VERSION_FILE, "w") as f:
                            f.write(latest_tag)
                            
                        messagebox.showinfo("Update Success", "Update installed successfully. The application will now start.")
                    else:
                        messagebox.showerror("Update Failed", "Failed to download the update.")
                else:
                    messagebox.showerror("Update Error", f"Could not find {APP_EXECUTABLE_NAME} in the release.")
            
            root.destroy()
            
    launch_app()

if __name__ == "__main__":
    main()
