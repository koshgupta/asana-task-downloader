import tkinter as tk
from tkinter import ttk
import threading
import zipfile
import platform
import shutil
import subprocess
from pathlib import Path
import asana
import requests
import os
from asana.rest import ApiException
from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import re

# --- UI for entering Asana credentials (token & project ID) ---
class CredentialsWindow:
    def __init__(self):
        self.token = None
        self.project_id = None

        # Create main Tk window
        self.root = tk.Tk()
        self.root.title("Asana Credentials")

        # Center window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width, window_height = 300, 150
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Frame for padding and layout
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # API Token input
        ttk.Label(main_frame, text="API Token:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.token_entry = ttk.Entry(main_frame, width=40)
        self.token_entry.grid(row=0, column=1, padx=5, pady=5)

        # Project ID input
        ttk.Label(main_frame, text="Project ID:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.project_entry = ttk.Entry(main_frame, width=40)
        self.project_entry.grid(row=1, column=1, padx=5, pady=5)

        # Button to submit and close window
        submit_button = ttk.Button(main_frame, text="Start Download", command=self.submit)
        submit_button.grid(row=2, column=0, columnspan=2, pady=20)

        # Make columns expand properly
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def submit(self):
        # Save entered values and close UI
        self.token = self.token_entry.get()
        self.project_id = self.project_entry.get()
        self.root.destroy()

    def get_credentials(self):
        # Run the Tk event loop until window is closed, then return inputs
        self.root.mainloop()
        return self.token, self.project_id


# --- Simple loading window with indeterminate progress bar ---
class LoadingWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Downloading Files")

        # Center window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width, window_height = 200, 100
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Keep window on top, platform-specific tweaks
        if platform.system() == "Darwin":  # macOS
            self.root.attributes('-topmost', True)
            self.root.createcommand('tk::mac::ReopenApplication', self.root.lift)
        else:  # Windows/Linux
            self.root.attributes('-topmost', True)
            self.root.overrideredirect(True)

        # Label and progress bar
        label = tk.Label(self.root, text="Downloading files...\nPlease wait", pady=10)
        label.pack()
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=150)
        self.progress.pack(pady=10)
        self.progress.start(10)  # animate

    def close(self):
        # Destroy window when done
        self.root.destroy()


def run_with_loading_window(main_function):
    """
    Wraps a long-running task so that a loading window is shown
    until the task completes.
    """
    loading_window = LoadingWindow()

    def task():
        try:
            main_function()
        finally:
            # Ensure the loading window closes on completion
            loading_window.root.after(0, loading_window.close)

    # Run the task in a background thread
    thread = threading.Thread(target=task)
    thread.start()

    # Start the loading UI loop
    loading_window.root.mainloop()


# --- Asana API helper functions ---

def get_tasks(tasks_api_instance, project_gid):
    """
    Retrieve all tasks in the project, including each task’s name.
    Returns a list of dicts: [{'gid': '123', 'name': 'My Task'}, …].
    """
    try:
        # ask Asana to also return the 'name' field
        opts = {'opt_fields': 'gid,name'}
        api_response = tasks_api_instance.get_tasks_for_project(project_gid, opts)

        tasks = []
        for data in api_response:
            tasks.append({
                'gid': data['gid'],
                'name': data['name'] or f"task-{data['gid']}"
            })
        return tasks

    except ApiException as e:
        print(f"Exception when fetching tasks: {e}")
        return []


# def get_attachment_gids(attachments_api_instance, tasks):
#     """
#     Given a list of task GIDs, retrieve all attachment GIDs.
#     """
#     list_of_attachments = []
#     try:
#         opts = {}
#         for task_gid in tasks:
#             api_response = attachments_api_instance.get_attachments_for_object(task_gid, opts)
#             for data in api_response:
#                 list_of_attachments.append(data['gid'])
#         return list_of_attachments
#     except ApiException as e:
#         print("Exception when calling AttachmentsApi->get_attachments: %s\n" % e)


def sanitize_filename(name):
    # remove or replace characters invalid in filenames
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()

def get_attachments_by_task(attachments_api_instance, tasks):
    """
    For each task dict {'gid','name'}, fetch its attachments and
    return { sanitized_task_name: [(filename, url), …], … }.
    """
    attachments_by_task = {}
    for task in tasks:
        gid = task['gid']
        raw_name = task['name']
        folder_name = sanitize_filename(raw_name)
        attachments_by_task[folder_name] = []

        try:
            # list all attachments on this task
            api_list = attachments_api_instance.get_attachments_for_object(gid, {})
            for att in api_list:
                detail = attachments_api_instance.get_attachment(att['gid'], {})
                fname = detail['name']
                url = detail['download_url']
                attachments_by_task[folder_name].append((fname, url))

        except ApiException as e:
            print(f"Error fetching attachments for task {gid}: {e}")

    return attachments_by_task


# --- File download & management functions ---

# def download_file(filename, url):
#     """
#     Download a single file from `url` into a dated folder.
#     """
#     date = datetime.date.today()
#     directory = f"usstm_asana_files-{date}"
#     os.makedirs(directory, exist_ok=True)

#     response = requests.get(url, stream=True)
#     if response.status_code == 200:
#         with open(f"{directory}/{filename}", "wb") as file:
#             for chunk in response.iter_content(chunk_size=8192):
#                 file.write(chunk)
#         print(f"Downloaded: {filename}")
#     else:
#         print(f"Failed to download {filename}: {response.status_code}")


def move_file_to_downloads(filepath):
    """
    Move the given file to the user's Downloads folder.
    """
    downloads_folder = str(Path.home() / "Downloads")
    os.makedirs(downloads_folder, exist_ok=True)
    new_filepath = os.path.join(downloads_folder, os.path.basename(filepath))
    shutil.move(filepath, new_filepath)
    print(f"File moved to: {new_filepath}")
    return new_filepath


def show_file_in_explorer(filepath):
    """
    Open the file location in the OS file explorer.
    """
    if platform.system() == "Windows":
        subprocess.run(["explorer", "/select,", os.path.abspath(filepath)], check=False)
    elif platform.system() == "Darwin":
        subprocess.run(["open", "-R", os.path.abspath(filepath)], check=False)
    else:
        print(f"File is saved at: {filepath}. Please open it manually.")


def download_files_grouped(attachments_by_task):
    """
    attachments_by_task: { 'Task A': [(file1, url1),…], 'Task B': […], … }
    Creates folder usstm_asana_files-<date>/Task A/... etc.
    Zips the entire dated folder, moves zip to Downloads, cleans up.
    """
    today = datetime.date.today()
    base_dir = f"usstm_asana_files-{today}"
    os.makedirs(base_dir, exist_ok=True)

    # download each file into its task’s folder
    for task_name, file_list in attachments_by_task.items():
        task_dir = os.path.join(base_dir, task_name)
        os.makedirs(task_dir, exist_ok=True)

        for filename, url in file_list:
            # stream download into the right subfolder
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                path = os.path.join(task_dir, filename)
                with open(path, 'wb') as f:
                    for chunk in response.iter_content(8192):
                        f.write(chunk)
                print(f"Downloaded {filename} into {task_name}/")
            else:
                print(f"Failed {filename}: {response.status_code}")

    # zip up the whole dated folder
    zip_path = f"{base_dir}.zip"
    with zipfile.ZipFile(zip_path, 'w') as z:
        for root, _, files in os.walk(base_dir):
            for fn in files:
                abs = os.path.join(root, fn)
                arc = os.path.relpath(abs, start=base_dir)
                z.write(abs, arc)

    print(f"Zipped to {zip_path}")
    # move & reveal as before
    dl = move_file_to_downloads(zip_path)
    show_file_in_explorer(dl)
    shutil.rmtree(base_dir)


# --- Main script entry point ---
if __name__ == "__main__":
    # Prompt user for Asana token & project ID
    creds_window = CredentialsWindow()
    token, project_gid = creds_window.get_credentials()

    # Only run if both values were provided
    if token and project_gid:
        # Initialize Asana client
        configuration = asana.Configuration()
        configuration.access_token = token
        api_client = asana.ApiClient(configuration)
        tasks_api_instance = asana.TasksApi(api_client)
        attachments_api_instance = asana.AttachmentsApi(api_client)

        # Define main download workflow
        def main_task():
            # 1) get tasks with names
            tasks = get_tasks(tasks_api_instance, project_gid)
        
            # 2) build a per-task attachment map
            attachments_by_task = get_attachments_by_task(attachments_api_instance, tasks)
        
            # 3) download into per-task folders and zip
            download_files_grouped(attachments_by_task)

        # Run the workflow with a loading/progress window
        run_with_loading_window(main_task)
