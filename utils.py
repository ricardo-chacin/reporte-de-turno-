from PIL import Image, ImageTk
import tkinter as tk
from pathlib import Path

def cargar_logo(parent):
    import tkinter as tk
    from PIL import Image, ImageTk
    from pathlib import Path

    base_path = Path(__file__).resolve().parent / "assets"
    logos = [
        ("logo_cerveceria.png", (50, 50)),
        ("Logo_nuevo.png",    (80, 36)),
    ]

    for filename, size in logos:
        try:
            logo_path = base_path / filename
            img = Image.open(logo_path).resize(size, Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(parent, image=photo, bg='#0d1117')
            lbl.image = photo
            lbl.pack(side='right', padx=(4, 6), pady=6)  # pady para centrar verticalmente
        except Exception as e:
            print(f"Error cargando {filename}: {e}")