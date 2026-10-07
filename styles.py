# styles.py - Tema oscuro minimalista
from ttkthemes import ThemedStyle

def apply_styles(root):
    style = ThemedStyle(root)
    style.set_theme("equilux")  # Tema oscuro profundo y limpio

    # Colores base oscuros
    style.configure('TFrame', background='#0d1117')
    style.configure('TLabel', background='#0d1117', foreground='#e9ecef', font=('Helvetica', 12))
    style.configure('TEntry', fieldbackground='#161b22', foreground='#e9ecef', font=('Helvetica', 12))
    style.configure('TCombobox', fieldbackground='#161b22', foreground='#e9ecef', font=('Helvetica', 12))
    style.configure('TButton', background='#212529', foreground='#e9ecef', font=('Helvetica', 14, 'bold'))
    style.map('TButton', background=[('active', '#30363d')])

    # Tarjetas oscuras con borde
    style.configure('Card.TFrame', background='#161b22', relief='groove', borderwidth=2)
    style.configure('Card.TLabel', background='#161b22', foreground='#e9ecef', font=('Helvetica', 14, 'bold'))

    return style