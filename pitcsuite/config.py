import os
import sys

def config_path():
    """Always resolves config.ini in the user's AppData directory to avoid permission issues."""
    appdata = os.getenv('APPDATA')
    folder = os.path.join(appdata, 'PITCSuite')
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, 'config.ini')
