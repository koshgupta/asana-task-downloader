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
import datetime
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from asana.rest import ApiException

# --- UI for entering Asana credentials
class CredentialsWindow:
    def __init__(self):
        self.token = None
        self.project_id = None
        self.root = tk.Tk()
        self.root.title("Asana Credentials")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width, window_height = 300, 150
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        ttk.Label(main_frame, text="API Token:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.token_entry = ttk.Entry(main_frame, width=40)
        self.token_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(main_frame, text="Project ID:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.project_entry = ttk.Entry(main_frame, width=40)
        self.project_entry.grid(row=1, column=1, padx=5, pady=5)
        submit_button = ttk.Button(main_frame, text="Start Download", command=self.submit)
        submit_button.grid(row=2, column=0, columnspan=2, pady=20)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def submit(self):
        self.token = self.token_entry.get()
        self.project_id = self.project_entry.get()
        self.root.destroy()

    def get_credentials(self):
        self.root.mainloop()
        return self.token, self.project_id

# --- Loading window
class LoadingWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Downloading Files")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width, window_height = 200, 100
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        if platform.system() == "Darwin":
            self.root.attributes('-topmost', True)
            self.root.createcommand('tk::mac::ReopenApplication', self.root.lift)
        else:
            self.root.attributes('-topmost', True)
            self.root.overrideredirect(True)
        label = tk.Label(self.root, text="Downloading files...\nPlease wait", pady=10)
        label.pack()
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=150)
        self.progress.pack(pady=10)
        self.progress.start(10)

    def close(self):
        self.root.destroy()

def run_with_loading_window(main_function):
    loading_window = LoadingWindow()
    def task():
        try:
            main_function()
        finally:
            loading_window.root.after(0, loading_window.close)
    thread = threading.Thread(target=task)
    thread.start()
    loading_window.root.mainloop()

# --- Helpers ---
def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()

# --- 1) Session for HTTP pooling ---
http_session = requests.Session()

# --- 2) Parallel Asana API for attachments ---
def get_tasks(tasks_api, project_gid):
    try:
        opts = {'opt_fields': 'gid,name'}
        api_response = tasks_api.get_tasks_for_project(project_gid, opts)
        return [
            {'gid': d['gid'], 'name': d['name'] or f"task-{d['gid']}"}
            for d in api_response
        ]
    except ApiException as e:
        print(f"Exception when fetching tasks: {e}")
        return []

def get_attachments_by_task(attachments_api, tasks):
    # Stage 1: fetch attachment lists in parallel
    def fetch_list(task):
        try:
            lst = attachments_api.get_attachments_for_object(task['gid'], {})
        except ApiException as e:
            print(f"Error listing attachments for {task['gid']}: {e}")
            lst = []
        return task, lst

    with ThreadPoolExecutor() as exe:
        lists = list(exe.map(fetch_list, tasks))

    # Stage 2: fetch details in parallel
    def fetch_detail(args):
        task, att = args
        name = sanitize_filename(task['name'])
        try:
            detail = attachments_api.get_attachment(att['gid'], {})
            return name, (detail['name'], detail['download_url'])
        except ApiException as e:
            print(f"Error fetching attachment {att['gid']}: {e}")
            return name, None

    jobs = []
    for task, att_list in lists:
        for att in att_list:
            jobs.append((task, att))

    attachments_by_task = {sanitize_filename(t['name']): [] for t in tasks}
    with ThreadPoolExecutor() as exe:
        for task_name, fileinfo in exe.map(fetch_detail, jobs):
            if fileinfo:
                attachments_by_task[task_name].append(fileinfo)

    return attachments_by_task

# --- 3) Parallelized download & 5) fast zip ---
def download_one(task_name, filename, url, base_dir):
    task_dir = base_dir / task_name
    path = task_dir / filename
    # stream download
    resp = http_session.get(url, stream=True)
    resp.raise_for_status()
    with open(path, 'wb') as f:
        for chunk in resp.iter_content(8192):
            f.write(chunk)
    return task_name, filename

def download_files_grouped(attachments_by_task):
    today = datetime.date.today()
    base_dir = Path(f"usstm_asana_files-{today}")
    base_dir.mkdir(exist_ok=True)

    # prepare directories & jobs
    jobs = []
    for task_name, files in attachments_by_task.items():
        (base_dir / task_name).mkdir(exist_ok=True)
        for fname, url in files:
            jobs.append((task_name, fname, url, base_dir))

    # parallel downloads
    with ThreadPoolExecutor(max_workers=8) as exe:
        futures = [exe.submit(download_one, *job) for job in jobs]
        for fut in as_completed(futures):
            try:
                tn, fn = fut.result()
                print(f"Downloaded {fn} into {tn}/")
            except Exception as e:
                print("Download error:", e)

    # fast C-level zip
    shutil.make_archive(str(base_dir), 'zip', root_dir=base_dir)

    # move to Downloads & reveal
    zip_path = f"{base_dir}.zip"
    downloads = Path.home() / "Downloads"
    downloads.mkdir(exist_ok=True)
    dest = downloads / Path(zip_path).name
    shutil.move(zip_path, dest)
    if platform.system() == "Windows":
        subprocess.run(["explorer", "/select,", str(dest)], check=False)
    elif platform.system() == "Darwin":
        subprocess.run(["open", "-R", str(dest)], check=False)
    else:
        print(f"Saved to {dest}")

    # clean up
    shutil.rmtree(base_dir)

# --- Main ---
if __name__ == "__main__":
    creds = CredentialsWindow()
    token, project_gid = creds.get_credentials()

    if token and project_gid:
        # Asana client
        configuration = asana.Configuration()
        configuration.access_token = token
        api_client = asana.ApiClient(configuration)
        tasks_api = asana.TasksApi(api_client)
        attachments_api = asana.AttachmentsApi(api_client)

        def main_task():
            # Profile if desired: import time; t0=time.perf_counter()
            tasks = get_tasks(tasks_api, project_gid)
            attachments = get_attachments_by_task(attachments_api, tasks)
            download_files_grouped(attachments)

        run_with_loading_window(main_task)