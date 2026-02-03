import tkinter as tk
from tkinter import ttk
import threading
import time
from PIL import Image, ImageTk
import os

def get_version():
    try:
        with open('VERSION.txt', encoding='utf-8') as f:
            return f.read().strip()
    except Exception:
        return "버전정보없음"

# Remove the loading screen feature entirely
# Delete the LoadingScreen class and show_loading_with_callback function 