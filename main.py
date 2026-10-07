import tkinter as tk
from ui import create_ui

FORMATS = {
    "Formato 175":       {"factor": 0.00175, "vel": 72000},
    "Formato 250": {"factor": 0.00250, "vel": 65000},
    "Formato 250": {"factor": 0.00250, "vel": 65000},
    "Formato 330": {"factor": 0.00330, "vel": 72000},
    "Formato 1000": {"factor": 0.01000, "vel": 19000},
}

if __name__ == "__main__":
    root = tk.Tk()
    create_ui(root, FORMATS)
    root.mainloop()