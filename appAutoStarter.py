import os
import win32com.client

# Path to your application executable
app_path = r"C:\Users\Bansal Com\AppData\Local\Postman\Postman.exe"  # Update this path to your app's executable
startup_folder = os.path.join(os.getenv("APPDATA"), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
shortcut_path = os.path.join(startup_folder, "Postman.lnk")  # Name of the shortcut

def create_shortcut():
    try:
        # Create a WScript.Shell object
        shell = win32com.client.Dispatch("WScript.Shell")
        
        # Create a shortcut object
        shortcut = shell.CreateShortcut(shortcut_path)
        
        # Set the target path to the application executable
        shortcut.TargetPath = app_path
        
        # Optionally set the working directory
        shortcut.WorkingDirectory = os.path.dirname(app_path)
        
        # Save the shortcut
        shortcut.save()
        
        print(f"Shortcut created at {shortcut_path}. Postman will start automatically on boot.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    create_shortcut()
