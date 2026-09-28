import os
import sys
import subprocess
import requests
import re
import threading
from packaging import version
from PySide6.QtWidgets import QMessageBox, QProgressDialog, QApplication
from PySide6.QtCore import QThread, Signal, Qt, QObject

GDRIVE_VERSION_FILE_ID = "1gmIKNgsj2hdYrrQtiwXHyQcXzCDR-ZZb" 
GDRIVE_SETUP_FILE_ID = "1wghUxDo3ezed5szbpXeWtWlyof4ykBIu"

def get_current_version():
    return "v2.8.83"

def get_gdrive_url_and_token(file_id):
    url = "https://docs.google.com/uc?export=download"
    session = requests.Session()
    response = session.get(url, params={'id': file_id}, stream=True)
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break
    if not token:
        match = re.search(r'confirm=([a-zA-Z0-9_-]+)', response.text)
        if match:
            token = match.group(1)
            
    final_url = f"{url}&id={file_id}"
    if token:
        final_url += f"&confirm={token}"
    return final_url

class DownloadThread(QThread):
    progress = Signal(int, int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, file_id, dest):
        super().__init__()
        self.file_id = file_id
        self.dest = dest

    def run(self):
        try:
            url = "https://docs.google.com/uc?export=download"
            session = requests.Session()
            response = session.get(url, params={'id': self.file_id}, stream=True)
            
            if "text/html" in response.headers.get("Content-Type", ""):
                text = response.text
                match = re.search(r'name="uuid" value="([^"]+)"', text)
                if match:
                    uuid = match.group(1)
                    final_url = f"https://drive.usercontent.google.com/download?id={self.file_id}&export=download&confirm=t&uuid={uuid}"
                    response = session.get(final_url, stream=True)
                else:
                    token = None
                    for key, value in response.cookies.items():
                        if key.startswith('download_warning'):
                            token = value
                            break
                    if not token:
                        match2 = re.search(r'confirm=([a-zA-Z0-9_-]+)', text)
                        if match2:
                            token = match2.group(1)
                    final_url = f"{url}&id={self.file_id}"
                    if token:
                        final_url += f"&confirm={token}"
                    response = session.get(final_url, stream=True)

            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            
            downloaded = 0
            with open(self.dest, "wb") as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)
                        downloaded += len(chunk)
                        self.progress.emit(downloaded, total_size)
            self.finished.emit(self.dest)
        except Exception as e:
            self.error.emit(str(e))

class UpdateChecker(QThread):
    result = Signal(bool, str) # update_available, version_string

    def run(self):
        try:
            version_url = get_gdrive_url_and_token(GDRIVE_VERSION_FILE_ID)
            v_response = requests.get(version_url, timeout=10)
            v_response.raise_for_status()
            latest_tag = v_response.text.strip()
            
            if latest_tag and latest_tag.startswith("v"):
                v_latest = version.parse(latest_tag.lstrip('v'))
                v_current = version.parse(get_current_version().lstrip('v'))
                if v_latest > v_current:
                    self.result.emit(True, latest_tag)
                    return
            self.result.emit(False, latest_tag)
        except Exception as e:
            self.result.emit(False, "")

def check_for_updates(parent_window, silent=False):
    checker = UpdateChecker(parent_window)
    parent_window._update_checker = checker # Prevent GC
    
    def on_result(has_update, latest_tag):
        if has_update:
            if silent:
                if hasattr(parent_window, 'notify_update_available'):
                    parent_window.notify_update_available()
            else:
                reply = QMessageBox.question(
                    parent_window,
                    "Update Available",
                    f"A new version ({latest_tag}) is available!\n\nYou are on {get_current_version()}.\n\nWould you like to download and install it now?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    start_download(parent_window)
        else:
            if not silent:
                if latest_tag == "":
                    QMessageBox.warning(parent_window, "Update Failed", "Could not check for updates.")
                else:
                    QMessageBox.information(parent_window, "Up to Date", f"You are running the latest version ({get_current_version()}).")
    
    checker.result.connect(on_result)
    checker.start()

def start_download(parent_window):
    progress_dialog = QProgressDialog("Downloading Update...", "Cancel", 0, 100, parent_window)
    progress_dialog.setWindowTitle("Updating PITC Suite")
    progress_dialog.setWindowModality(Qt.WindowModal)
    progress_dialog.setAutoClose(True)
    progress_dialog.setAutoReset(True)
    progress_dialog.resize(400, 100)

    dest_path = os.path.join(os.environ["TEMP"], "PITCSuite_Setup.exe")
    
    downloader = DownloadThread(GDRIVE_SETUP_FILE_ID, dest_path)
    parent_window._downloader = downloader # PREVENT GARBAGE COLLECTION CRASH
    
    def on_progress(downloaded, total):
        if total > 0:
            percent = int((downloaded / total) * 100)
            progress_dialog.setValue(percent)
        else:
            progress_dialog.setValue(0)
            progress_dialog.setLabelText(f"Downloading... {downloaded // 1024 // 1024} MB")

    def on_finished(path):
        progress_dialog.close()
        
        msg = QMessageBox(parent_window)
        msg.setWindowTitle("Application Update")
        msg.setText("The update is ready to install.")
        msg.setInformativeText("The application will restart and install the update silently in the background.")
        msg.addButton("Restart application", QMessageBox.AcceptRole)
        msg.exec()
        
        # CRITICAL FIX for PyInstaller subprocess crash:
        # We MUST strip PyInstaller's _PYI_* environment variables, otherwise the newly launched 
        # app inherits them and gets fatally confused, trying to use the old deleted Temp folder!
        env = os.environ.copy()
        keys_to_remove = [k for k in env if k.startswith('_PYI_') or k.startswith('TCL_') or k.startswith('TK_')]
        for k in keys_to_remove:
            env.pop(k, None)
            
        if hasattr(sys, '_MEIPASS'):
            meipass = sys._MEIPASS.lower()
            path_dirs = env.get('PATH', '').split(os.pathsep)
            path_dirs = [p for p in path_dirs if p.lower().rstrip(r'\/') != meipass]
            env['PATH'] = os.pathsep.join(path_dirs)
            
        # Launch Setup.exe directly (0x00000008 is DETACHED_PROCESS)
        # We don't use a .bat file so the user can see the progress bar properly.
        subprocess.Popen([path, "/SILENT", "/FORCECLOSEAPPLICATIONS", "/AUTOUPDATE=1"], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
        
        QApplication.instance().quit()
        sys.exit(0)

    def on_error(err):
        progress_dialog.close()
        QMessageBox.critical(parent_window, "Download Failed", f"An error occurred while downloading:\n{err}")

    downloader.progress.connect(on_progress)
    downloader.finished.connect(on_finished)
    downloader.error.connect(on_error)
    
    progress_dialog.canceled.connect(downloader.terminate)
    
    downloader.start()
    progress_dialog.show()







