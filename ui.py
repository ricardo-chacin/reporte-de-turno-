import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import Calendar
import pandas as pd
import os
import sys

# Windows 100% escala
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from styles import apply_styles
from utils import cargar_logo
from logic import calcular_resultados
from storage import guardar_en_excel


def get_color(kpi, value_str):
    try:
        v = float(value_str.rstrip('%'))
        if kpi in ["GLY", "OSE"]:
            return '#2ecc71' if v >= 85 else '#f1c40f' if v >= 75 else '#e74c3c'
        elif kpi == "LEF":
            return '#2ecc71' if v >= 90 else '#f1c40f' if v >= 80 else '#e74c3c'
        elif kpi in ["Merma", "Merma F1", "Merma F2"]:
            return '#2ecc71' if v <= 1.0 else '#f1c40f' if v <= 2.5 else '#e74c3c'
        elif kpi == "Cumplimiento":
            return '#2ecc71' if v >= 100 else '#f1c40f' if v >= 90 else '#e74c3c'
        elif kpi == "Agua":
            return '#2ecc71' if v <= 1.5 else '#f1c40f' if v <= 2.0 else '#e74c3c'
        elif kpi == "CO2":
            return '#2ecc71' if v <= 2.0 else '#f1c40f' if v <= 2.5 else '#e74c3c'
        elif kpi == "Vapor":
            return '#2ecc71' if v <= 40  else '#f1c40f' if v <= 50  else '#e74c3c'
    except Exception:
        pass
    return '#c9d1d9'


def kpi_card(parent, kpi, value, row, col):
    card = tk.Frame(parent, bg='#161b22',
                    highlightbackground='#30363d', highlightthickness=1)
    card.grid(row=row, column=col, padx=5, pady=5, ipadx=10, ipady=8, sticky='nsew')
    tk.Label(card, text=kpi,  font=('Helvetica', 9),  fg='#8b949e', bg='#161b22').pack(anchor='center')
    tk.Label(card, text=value, font=('Helvetica', 16, 'bold'),
             fg=get_color(kpi, value), bg='#161b22').pack(anchor='center')


# ════════════════════════════════════════════════════════════════════════════
def create_ui(root, formats):
    apply_styles(root)
    try:
        root.tk.call('tk', 'scaling', 1.0)
    except Exception:
        pass
    root.title("Reporte de Turno")
    root.geometry("1200x820")
    root.configure(bg='#0d1117')
    root.minsize(1100, 700)

    productos = list(formats.keys())

    # ── Header compacto ───────────────────────────────────────────────────────
    header = tk.Frame(root, bg='#0d1117', height=48)
    header.pack(fill='x', padx=14, pady=(6, 4))
    header.pack_propagate(False)

    ttk.Label(header, text="Reporte de Turno",
              font=('Helvetica', 16, 'bold'), foreground='#C8102E').pack(side='left', pady=6)
    cargar_logo(header)

    tk.Frame(root, bg='#30363d', height=1).pack(fill='x', padx=14)

    # Frames principales
    frame1 = tk.Frame(root, bg='#0d1117')
    frame1.pack(fill='both', expand=True)
    frame2 = tk.Frame(root, bg='#0d1117')
    frame3 = tk.Frame(root, bg='#0d1117')

    # ════════════════════════════════════════════════════════════════════════
    # FRAME 1
    # ════════════════════════════════════════════════════════════════════════
    hubo_cambio = [False]

    # Toolbar
    toolbar = tk.Frame(frame1, bg='#0d1117')
    toolbar.pack(fill='x', padx=14, pady=(8, 4))
    ttk.Label(toolbar, text="ENTRADAS DEL TURNO",
              font=('Helvetica', 12, 'bold'), foreground='#C8102E').pack(side='left')
    ttk.Label(toolbar, text="Turno:", foreground='#c9d1d9',
              font=('Helvetica', 11)).pack(side='left', padx=(20, 4))
    turno_var = tk.StringVar(value="T1")
    ttk.Combobox(toolbar, textvariable=turno_var,
                 values=["T1", "T2", "T3", "DIA"], width=6).pack(side='left')

    # Botón historial en toolbar
    def abrir_historial():
        frame1.pack_forget()
        frame2.pack_forget()
        construir_historial()
        frame3.pack(fill='both', expand=True)

    ttk.Button(toolbar, text="📋  HISTORIAL", width=14,
               command=abrir_historial).pack(side='right')

    tk.Frame(frame1, bg='#30363d', height=1).pack(fill='x', padx=14, pady=(4, 0))

    # ── Cuerpo 3 columnas ─────────────────────────────────────────────────────
    body = tk.Frame(frame1, bg='#0d1117')
    body.pack(fill='both', expand=True)
    body.columnconfigure(0, weight=1)
    body.columnconfigure(1, weight=1)
    body.columnconfigure(2, weight=1)
    body.rowconfigure(0, weight=1)

    col1 = tk.Frame(body, bg='#0d1117', highlightbackground='#21262d', highlightthickness=1)
    col1.grid(row=0, column=0, sticky='nsew')
    col2 = tk.Frame(body, bg='#0d1117', highlightbackground='#21262d', highlightthickness=1)
    col2.grid(row=0, column=1, sticky='nsew')
    col3 = tk.Frame(body, bg='#0d1117', highlightbackground='#21262d', highlightthickness=1)
    col3.grid(row=0, column=2, sticky='nsew')

    # ── Helpers ───────────────────────────────────────────────────────────────
    def col_title(parent, text, color='#C8102E'):
        tk.Label(parent, text=text, font=('Helvetica', 12, 'bold'),
                 fg=color, bg='#0d1117').pack(anchor='w', padx=10, pady=(6, 2))
        tk.Frame(parent, bg='#21262d', height=1).pack(fill='x', padx=10, pady=(0, 4))

    def sub_label(parent, text):
        tk.Label(parent, text=text, font=('Helvetica', 11, 'bold'),
                 fg='#C8102E', bg='#0d1117').pack(anchor='w', padx=10, pady=(6, 1))

    def field_row(parent, label, width=14, unit=""):
        row = tk.Frame(parent, bg='#0d1117')
        row.pack(fill='x', padx=10, pady=2)
        tk.Label(row, text=label, font=('Helvetica', 11), fg='#8b949e',
                 bg='#0d1117', width=18, anchor='w').pack(side='left')
        e = ttk.Entry(row, width=width)
        e.pack(side='left')
        if unit:
            tk.Label(row, text=unit, font=('Helvetica', 9),
                     fg='#58a6ff', bg='#0d1117').pack(side='left', padx=(5, 0))
        return e

    # ════════════════════════════════════════════════════════════════════════
    # COLUMNA 1 — Formato (tabs F1 / F2)
    # ════════════════════════════════════════════════════════════════════════
    tabs_frame = tk.Frame(col1, bg='#0d1117')
    tabs_frame.pack(fill='x', padx=10, pady=(8, 0))

    tab_f1_btn = tk.Label(tabs_frame, text="Formato 1", font=('Helvetica', 9, 'bold'),
                          fg='#58a6ff', bg='#21262d', padx=12, pady=4,
                          relief='flat', cursor='hand2')
    tab_f1_btn.pack(side='left')
    tab_f2_btn = tk.Label(tabs_frame, text="Formato 2", font=('Helvetica', 9, 'bold'),
                          fg='#8b949e', bg='#0d1117', padx=12, pady=4,
                          relief='flat', cursor='hand2')
    tab_f2_btn.pack(side='left')

    tk.Frame(col1, bg='#21262d', height=1).pack(fill='x', padx=10, pady=(2, 4))

    campos_container = tk.Frame(col1, bg='#0d1117')
    campos_container.pack(fill='both', expand=True)

    def build_formato_fields(parent):
        entries = {}
        var_prod = tk.StringVar(value=productos[1])
        row = tk.Frame(parent, bg='#0d1117')
        row.pack(fill='x', padx=10, pady=2)
        tk.Label(row, text="Producto:", font=('Helvetica', 11), fg='#8b949e',
                 bg='#0d1117', width=18, anchor='w').pack(side='left')
        cb = ttk.Combobox(row, textvariable=var_prod, values=productos, width=17)
        cb.pack(side='right')
        entries["Producto:"] = var_prod

        # (label, unidad)
        simple_fields = [
            ("Prog. (botellas):",   "bot"),
            ("Env. (botellas):",    "bot"),
            ("Rechazo Rotuladora:", "bot"),
            ("Explosiones (bot):",  "bot"),
            ("Rotura (bot):",       "bot"),
            ("Másico A Inicial:",   "HL"),
            ("Másico A Final:",     "HL"),
            ("Másico B Inicial:",   "HL"),
            ("Másico B Final:",     "HL"),
        ]
        for lbl, unit in simple_fields:
            e = field_row(parent, lbl, unit=unit)
            entries[lbl] = e
        return entries

    panel_f1 = tk.Frame(campos_container, bg='#0d1117')
    panel_f1.pack(fill='both', expand=True)
    entries_f1 = build_formato_fields(panel_f1)

    panel_f2 = tk.Frame(campos_container, bg='#0d1117')
    entries_f2 = build_formato_fields(panel_f2)

    def show_tab(tab):
        if tab == 1:
            panel_f2.pack_forget()
            panel_f1.pack(fill='both', expand=True)
            tab_f1_btn.config(fg='#58a6ff', bg='#21262d')
            tab_f2_btn.config(fg='#8b949e', bg='#0d1117')
        else:
            panel_f1.pack_forget()
            panel_f2.pack(fill='both', expand=True)
            tab_f1_btn.config(fg='#8b949e', bg='#0d1117')
            tab_f2_btn.config(fg='#f1c40f', bg='#21262d')

    tab_f1_btn.bind('<Button-1>', lambda e: show_tab(1))
    tab_f2_btn.bind('<Button-1>', lambda e: show_tab(2))

    # ════════════════════════════════════════════════════════════════════════
    # COLUMNA 2 — Servicios + Fallas
    # ════════════════════════════════════════════════════════════════════════
    col_title(col2, "SERVICIOS Y TIEMPOS")

    entries_srv = {}
    srv_fields = [
        ("Min. Perdidos Ext:", "min_perd", "min"),
        ("NST Demanda (min):", "nst",      "min"),
        ("DPA (minutos):",     "dpa",      "min"),
        ("Consumo Agua:",      "agua",     "m³"),
        ("Consumo Vapor:",     "vapor",    "kg"),
        ("Consumo CO2:",       "co2",      "kg"),
    ]

    srv_grid = tk.Frame(col2, bg='#0d1117')
    srv_grid.pack(fill='x', padx=10, pady=(0, 4))
    srv_grid.columnconfigure(0, weight=1)
    srv_grid.columnconfigure(1, weight=1)

    for i, (label, key, unit) in enumerate(srv_fields):
        cell = tk.Frame(srv_grid, bg='#0d1117')
        cell.grid(row=i // 2, column=i % 2, padx=4, pady=3, sticky='w')
        tk.Label(cell, text=label, font=('Helvetica', 11),
                 fg='#8b949e', bg='#0d1117').pack(anchor='w')
        inp_row = tk.Frame(cell, bg='#0d1117')
        inp_row.pack(anchor='w')
        e = ttk.Entry(inp_row, width=12)
        e.pack(side='left')
        tk.Label(inp_row, text=unit, font=('Helvetica', 9),
                 fg='#58a6ff', bg='#0d1117').pack(side='left', padx=(5, 0))
        entries_srv[key] = e

    tk.Frame(col2, bg='#30363d', height=1).pack(fill='x', padx=10, pady=(6, 2))
    sub_label(col2, "FALLAS DEL TURNO")
    txt_fallas = tk.Text(col2, bg='#161b22', fg='#e9ecef', font=('Helvetica', 9),
                         insertbackground='white', relief='groove', borderwidth=1,
                         wrap='word', height=6)
    txt_fallas.pack(fill='both', expand=True, padx=10, pady=(2, 10))

    # ════════════════════════════════════════════════════════════════════════
    # COLUMNA 3 — Comentarios
    # ════════════════════════════════════════════════════════════════════════
    col_title(col3, "COMENTARIOS DEL TURNO")

    comentarios_entries = {}
    comentarios_entries["fallas"] = txt_fallas

    sub_label(col3, "ACTIVIDADES CORRECTIVAS")
    txt_correctivas = tk.Text(col3, bg='#161b22', fg='#e9ecef', font=('Helvetica', 9),
                              insertbackground='white', relief='groove', borderwidth=1,
                              wrap='word', height=4)
    txt_correctivas.pack(fill='x', padx=10, pady=(2, 4))
    comentarios_entries["correctivas"] = txt_correctivas

    sub_label(col3, "ANÁLISIS 5W")
    txt_5w = tk.Text(col3, bg='#161b22', fg='#e9ecef', font=('Helvetica', 9),
                     insertbackground='white', relief='groove', borderwidth=1,
                     wrap='word', height=4)
    txt_5w.pack(fill='x', padx=10, pady=(2, 4))
    comentarios_entries["cinco_w"] = txt_5w

    sub_label(col3, "AVISOS / NOVEDADES")
    avisos_lista = []
    av_lista_frame = tk.Frame(col3, bg='#0d1117')
    av_lista_frame.pack(fill='x', padx=10, pady=(2, 0))

    def agregar_aviso(num="", desc=""):
        fila = tk.Frame(av_lista_frame, bg='#0d1117')
        fila.pack(fill='x', pady=2)
        tk.Label(fila, text="N°:", fg='#8b949e', bg='#0d1117',
                 font=('Helvetica', 11)).pack(side='left')
        e_num = ttk.Entry(fila, width=7)
        e_num.pack(side='left', padx=(2, 6))
        if num: e_num.insert(0, num)
        tk.Label(fila, text="Desc:", fg='#8b949e', bg='#0d1117',
                 font=('Helvetica', 11)).pack(side='left')
        e_desc = ttk.Entry(fila, width=18)
        e_desc.pack(side='left', padx=(2, 4))
        if desc: e_desc.insert(0, desc)

        def eliminar():
            avisos_lista.remove((e_num, e_desc, fila))
            fila.destroy()

        lbl_x = tk.Label(fila, text="✕", fg='#e74c3c', bg='#0d1117',
                         font=('Helvetica', 10, 'bold'), cursor='hand2')
        lbl_x.pack(side='left')
        lbl_x.bind('<Button-1>', lambda e: eliminar())
        avisos_lista.append((e_num, e_desc, fila))

    agregar_aviso()
    ttk.Button(col3, text="+ Agregar aviso",
               command=agregar_aviso, width=16).pack(anchor='w', padx=10, pady=(4, 0))

    # ── Footer con botones ────────────────────────────────────────────────────
    tk.Frame(frame1, bg='#30363d', height=1).pack(fill='x', padx=14)
    btn_row = tk.Frame(frame1, bg='#0d1117')
    btn_row.pack(fill='x', padx=14, pady=8)

    def activar_f2():
        if not hubo_cambio[0]:
            hubo_cambio[0] = True
            show_tab(2)
            btn_cambio.config(text="✕  Quitar Formato 2")
        else:
            hubo_cambio[0] = False
            show_tab(1)
            btn_cambio.config(text="+  Cambio de Formato")

    btn_cambio = ttk.Button(btn_row, text="+  Cambio de Formato",
                            command=activar_f2, width=22)
    btn_cambio.pack(side='left', padx=(0, 10))
    ttk.Button(btn_row, text="CALCULAR TURNO", width=22,
               command=lambda: calcular()).pack(side='left')

    # ════════════════════════════════════════════════════════════════════════
    # CALCULAR
    # ════════════════════════════════════════════════════════════════════════
    def calcular():
        try:
            def fget(d, k):
                return float(d[k].get() or 0)
            def sget(key):
                return float(entries_srv[key].get() or 0)

            formato_f2_val = entries_f2["Producto:"].get() if hubo_cambio[0] else None

            resumen, det_f1, det_f2 = calcular_resultados(
                entries_f1["Producto:"].get(),
                fget(entries_f1, "Prog. (botellas):"),
                fget(entries_f1, "Env. (botellas):"),
                fget(entries_f1, "Rechazo Rotuladora:"),
                fget(entries_f1, "Explosiones (bot):"),
                fget(entries_f1, "Rotura (bot):"),
                fget(entries_f1, "Másico A Inicial:"),
                fget(entries_f1, "Másico A Final:"),
                fget(entries_f1, "Másico B Inicial:"),
                fget(entries_f1, "Másico B Final:"),
                sget("min_perd"), sget("nst"), sget("dpa"),
                sget("agua"), sget("vapor"), sget("co2"),
                formats,
                formato_f2_val,
                fget(entries_f2, "Prog. (botellas):")    if hubo_cambio[0] else 0,
                fget(entries_f2, "Env. (botellas):")     if hubo_cambio[0] else 0,
                fget(entries_f2, "Rechazo Rotuladora:")  if hubo_cambio[0] else 0,
                fget(entries_f2, "Explosiones (bot):")   if hubo_cambio[0] else 0,
                fget(entries_f2, "Rotura (bot):")        if hubo_cambio[0] else 0,
                fget(entries_f2, "Másico A Inicial:")    if hubo_cambio[0] else 0,
                fget(entries_f2, "Másico A Final:")      if hubo_cambio[0] else 0,
                fget(entries_f2, "Másico B Inicial:")    if hubo_cambio[0] else 0,
                fget(entries_f2, "Másico B Final:")      if hubo_cambio[0] else 0,
                turno=turno_var.get(),
            )

            comentarios = {c: w.get("1.0", "end-1c").strip()
                           for c, w in comentarios_entries.items()}
            comentarios["avisos"] = [
                {"numero": e_num.get().strip(), "descripcion": e_desc.get().strip()}
                for e_num, e_desc, _ in avisos_lista
                if e_num.get().strip() or e_desc.get().strip()
            ]

            datos_entrada = {
                "Turno": turno_var.get(),
                "Formato_F1": entries_f1["Producto:"].get(),
                "Formato_F2": formato_f2_val or "",
            }

            mostrar_resultado(resumen, det_f1, det_f2, comentarios,
                              datos_entrada, turno_var.get())

        except Exception as e:
            messagebox.showerror("Error", f"Error al calcular:\n{str(e)}")

    # ════════════════════════════════════════════════════════════════════════
    # MOSTRAR RESULTADO — reutilizable para calcular e historial
    # ════════════════════════════════════════════════════════════════════════
    def mostrar_resultado(resumen, det_f1, det_f2, comentarios,
                          datos_entrada=None, turno_label="", desde_historial=False):
        frame1.pack_forget()
        frame3.pack_forget()
        frame2.pack(fill='both', expand=True)
        for w in frame2.winfo_children():
            w.destroy()

        tbar = tk.Frame(frame2, bg='#0d1117')
        tbar.pack(fill='x', padx=14, pady=(8, 0))
        ttk.Label(tbar, text="REPORTE DE TURNO",
                  font=('Helvetica', 16, 'bold'), foreground='#C8102E').pack(side='left')
        ttk.Label(tbar, text=f"  {turno_label}",
                  font=('Helvetica', 12), foreground='#8b949e').pack(side='left', padx=6)

        def reset():
            hubo_cambio[0] = False
            show_tab(1)
            btn_cambio.config(text="+  Cambio de Formato")
            for txt in comentarios_entries.values():
                txt.delete("1.0", "end")
            for _, _, fila in avisos_lista[1:]:
                fila.destroy()
            del avisos_lista[1:]
            avisos_lista[0][0].delete(0, "end")
            avisos_lista[0][1].delete(0, "end")
            frame2.pack_forget()
            frame1.pack(fill='both', expand=True)

        def volver():
            frame2.pack_forget()
            if desde_historial:
                frame3.pack(fill='both', expand=True)
            else:
                frame1.pack(fill='both', expand=True)

        btn_bar = tk.Frame(tbar, bg='#0d1117')
        btn_bar.pack(side='right')
        if datos_entrada:
            ttk.Button(btn_bar, text="GUARDAR", width=14,
                       command=lambda: guardar_en_excel(datos_entrada, resumen,
                                                        comentarios, det_f1, det_f2)
                       ).pack(side='left', padx=(0, 8))
        ttk.Button(btn_bar,
                   text="EDITAR DATOS" if not desde_historial else "← VOLVER",
                   width=14, command=volver).pack(side='left', padx=(0, 4))
        if not desde_historial:
            ttk.Button(btn_bar, text="NUEVO TURNO", width=14,
                       command=reset).pack(side='left')

        tk.Frame(frame2, bg='#30363d', height=1).pack(fill='x', padx=14, pady=(6, 4))

        main_cols = tk.Frame(frame2, bg='#0d1117')
        main_cols.pack(fill='both', expand=True, padx=10, pady=(0, 6))
        main_cols.columnconfigure(0, weight=3)
        main_cols.columnconfigure(1, weight=2)
        main_cols.rowconfigure(0, weight=1)

        col_izq = tk.Frame(main_cols, bg='#0d1117')
        col_izq.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        col_der = tk.Frame(main_cols, bg='#0d1117')
        col_der.grid(row=0, column=1, sticky='nsew')

        ttk.Label(col_izq, text="RESUMEN GENERAL",
                  font=('Helvetica', 11, 'bold'), foreground='#C8102E'
                  ).pack(anchor='w', padx=8, pady=(4, 3))
        res_grid = ttk.Frame(col_izq)
        res_grid.pack(fill='x', padx=6, pady=(0, 4))
        kpis_orden = ["HL Envasados", "Cumplimiento", "GLY", "LEF",
                      "Merma", "OSE", "Agua", "CO2", "Vapor", "TE (EPT)"]
        for i, kpi in enumerate(kpis_orden):
            kpi_card(res_grid, kpi, resumen.get(kpi, "--"), i // 5, i % 5)
        for c in range(5):
            res_grid.columnconfigure(c, weight=1)

        tk.Frame(col_izq, bg='#30363d', height=1).pack(fill='x', padx=6, pady=(2, 4))
        ttk.Label(col_izq, text="DETALLE POR FORMATO",
                  font=('Helvetica', 11, 'bold'), foreground='#C8102E'
                  ).pack(anchor='w', padx=8, pady=(0, 3))

        def build_detalle_card(parent, detalle, color_titulo, tag):
            card = tk.Frame(parent, bg='#1c2333',
                            highlightbackground=color_titulo, highlightthickness=1)
            card.pack(fill='x', padx=6, pady=3, ipadx=8, ipady=5)
            hdr = tk.Frame(card, bg='#1c2333')
            hdr.pack(fill='x')
            tk.Label(hdr, text=f"{tag}  —  {detalle['Formato']}",
                     font=('Helvetica', 10, 'bold'),
                     bg='#1c2333', fg=color_titulo).pack(side='left')
            tk.Label(hdr, text=f"  Merma: {detalle['Merma']}",
                     font=('Helvetica', 10), bg='#1c2333',
                     fg=get_color('Merma', detalle['Merma'])).pack(side='left', padx=12)
            body_d = tk.Frame(card, bg='#1c2333')
            body_d.pack(fill='x', pady=(4, 0))
            for ki, (nombre, val) in enumerate([
                ("HL Prog.",  detalle.get("HL Programados", "--")),
                ("HL Env.",   detalle.get("HL Envasados",   "--")),
                ("Másico",    detalle.get("Vol. Másico",    "--")),
                ("Merma",     detalle.get("Merma",          "--")),
            ]):
                cel = tk.Frame(body_d, bg='#0d1117', padx=8, pady=4)
                cel.grid(row=0, column=ki, padx=4, sticky='nsew')
                tk.Label(cel, text=nombre, font=('Helvetica', 8),
                         fg='#8b949e', bg='#0d1117').pack(anchor='center')
                tk.Label(cel, text=val, font=('Helvetica', 13, 'bold'),
                         fg=get_color('Merma' if nombre == 'Merma' else '', val),
                         bg='#0d1117').pack(anchor='center')
            for c in range(4):
                body_d.columnconfigure(c, weight=1)

        build_detalle_card(col_izq, det_f1, '#58a6ff', "F1")
        if det_f2:
            build_detalle_card(col_izq, det_f2, '#f1c40f', "F2")

        # Gráfica
        ttk.Label(col_der, text="EFICIENCIAS CLAVE",
                  font=('Helvetica', 11, 'bold'), foreground='#8b949e'
                  ).pack(anchor='w', padx=8, pady=(4, 3))
        inds = ["GLY", "LEF", "OSE", "Merma"]
        vals, cols_graf = [], []
        for k in inds:
            try:    v = float(resumen.get(k, "0").rstrip('%'))
            except: v = 0
            vals.append(v)
            cols_graf.append(get_color(k, f"{v}"))

        fig = Figure(figsize=(5, 2.0), dpi=96, facecolor='#161b22')
        ax  = fig.add_subplot(111, facecolor='#161b22')
        fig.subplots_adjust(left=0.14, right=0.96, top=0.92, bottom=0.10)
        ax.barh(range(len(inds)), vals, color=cols_graf, height=0.45, edgecolor='#30363d')
        ax.set_yticks(range(len(inds)))
        ax.set_yticklabels(inds, color='white', fontsize=10, fontweight='bold')
        ax.tick_params(axis='x', colors='#8b949e', labelsize=8)
        ax.set_xlim(0, 115)
        ax.invert_yaxis()
        ax.axvline(85, color='#555', linestyle='--', linewidth=0.7)
        ax.axvline(90, color='#888', linestyle='--', linewidth=0.7)
        for spine in ['top', 'right']:
            ax.spines[spine].set_visible(False)
        for spine in ['left', 'bottom']:
            ax.spines[spine].set_color('#30363d')
        for i, (v, c) in enumerate(zip(vals, cols_graf)):
            ax.text(v + 0.5, i, f'{v:.1f}%', va='center', color=c, fontsize=9, fontweight='bold')

        graf_frame = ttk.Frame(col_der)
        graf_frame.pack(fill='x', padx=4, pady=(0, 6))
        fig_canvas = FigureCanvasTkAgg(fig, master=graf_frame)
        fig_canvas.draw()
        fig_canvas.get_tk_widget().config(highlightthickness=0, bd=0)
        fig_canvas.get_tk_widget().pack(fill='x')

        tk.Frame(col_der, bg='#30363d', height=1).pack(fill='x', padx=4, pady=(0, 4))
        ttk.Label(col_der, text="COMENTARIOS DEL TURNO",
                  font=('Helvetica', 11, 'bold'), foreground='#8b949e'
                  ).pack(anchor='w', padx=8, pady=(0, 3))
        for clave, titulo in [("fallas", "Fallas del turno"),
                               ("correctivas", "Actividades correctivas"),
                               ("cinco_w", "Análisis 5W")]:
            texto = comentarios.get(clave, "") or "Sin novedad"
            card_c = tk.Frame(col_der, bg='#161b22',
                              highlightbackground='#30363d', highlightthickness=1)
            card_c.pack(fill='x', padx=6, pady=2, ipadx=8, ipady=4)
            tk.Label(card_c, text=titulo, font=('Helvetica', 9, 'bold'),
                     bg='#161b22', fg='#C8102E').pack(anchor='w')
            tk.Label(card_c, text=texto, font=('Helvetica', 9),
                     bg='#161b22', fg='#c9d1d9',
                     wraplength=320, justify='left').pack(anchor='w', pady=(2, 0))

        avisos_list = comentarios.get("avisos", [])
        if avisos_list:
            av_card = tk.Frame(col_der, bg='#1c2333',
                               highlightbackground='#C8102E', highlightthickness=1)
            av_card.pack(fill='x', padx=6, pady=(4, 2), ipadx=8, ipady=4)
            tk.Label(av_card, text=f"AVISOS ({len(avisos_list)})",
                     font=('Helvetica', 9, 'bold'),
                     bg='#1c2333', fg='#C8102E').pack(anchor='w', pady=(0, 2))
            for av in avisos_list:
                av_r = tk.Frame(av_card, bg='#0d1117')
                av_r.pack(fill='x', pady=1, ipadx=4, ipady=2)
                tk.Label(av_r, text=f"N° {av.get('numero','')}" if av.get('numero') else "—",
                         font=('Helvetica', 9, 'bold'),
                         bg='#0d1117', fg='#f1c40f').pack(side='left', padx=(4, 10))
                tk.Label(av_r, text=av.get('descripcion') or "Sin descripción",
                         font=('Helvetica', 9), bg='#0d1117', fg='#c9d1d9',
                         wraplength=300, justify='left').pack(side='left')

    # ════════════════════════════════════════════════════════════════════════
    # FRAME 3 — HISTORIAL CON CALENDARIO
    # ════════════════════════════════════════════════════════════════════════
    def construir_historial():
        for w in frame3.winfo_children():
            w.destroy()

        tbar_h = tk.Frame(frame3, bg='#0d1117')
        tbar_h.pack(fill='x', padx=14, pady=(8, 0))
        ttk.Label(tbar_h, text="HISTORIAL DE TURNOS",
                  font=('Helvetica', 16, 'bold'), foreground='#C8102E').pack(side='left')
        ttk.Button(tbar_h, text="← VOLVER", width=12,
                   command=lambda: [frame3.pack_forget(),
                                    frame1.pack(fill='both', expand=True)]
                   ).pack(side='right')

        tk.Frame(frame3, bg='#30363d', height=1).pack(fill='x', padx=14, pady=(6, 4))

        hist_body = tk.Frame(frame3, bg='#0d1117')
        hist_body.pack(fill='both', expand=True, padx=14, pady=4)
        hist_body.columnconfigure(1, weight=1)
        hist_body.rowconfigure(0, weight=1)

        # ── Columna izquierda: calendario ─────────────────────────────────
        cal_frame = tk.Frame(hist_body, bg='#161b22',
                             highlightbackground='#30363d', highlightthickness=1)
        cal_frame.grid(row=0, column=0, sticky='n', padx=(0, 14), pady=4)

        ttk.Label(cal_frame, text="Selecciona una fecha",
                  font=('Helvetica', 10, 'bold'), foreground='#8b949e'
                  ).pack(padx=10, pady=(8, 4))

        cal = Calendar(cal_frame,
                       selectmode='day',
                       date_pattern='yyyy-mm-dd',
                       background='#161b22',
                       foreground='#e9ecef',
                       selectbackground='#C8102E',
                       selectforeground='white',
                       normalbackground='#161b22',
                       normalforeground='#e9ecef',
                       weekendbackground='#1c2333',
                       weekendforeground='#8b949e',
                       headersbackground='#0d1117',
                       headersforeground='#C8102E',
                       bordercolor='#30363d',
                       othermonthbackground='#0d1117',
                       othermonthforeground='#30363d',
                       font=('Helvetica', 9))
        cal.pack(padx=10, pady=4)

        ttk.Button(cal_frame, text="Ver reportes",
                   command=lambda: buscar_reportes(cal.get_date())
                   ).pack(pady=(4, 10))

        # ── Columna derecha: lista de reportes ────────────────────────────
        lista_frame = tk.Frame(hist_body, bg='#0d1117')
        lista_frame.grid(row=0, column=1, sticky='nsew', pady=4)
        lista_frame.rowconfigure(1, weight=1)
        lista_frame.columnconfigure(0, weight=1)

        lbl_fecha_sel = ttk.Label(lista_frame,
                                  text="Selecciona una fecha en el calendario",
                                  font=('Helvetica', 11), foreground='#8b949e')
        lbl_fecha_sel.grid(row=0, column=0, sticky='w', pady=(4, 8))

        canvas_wrap = tk.Frame(lista_frame, bg='#0d1117')
        canvas_wrap.grid(row=1, column=0, sticky='nsew')
        canvas_wrap.rowconfigure(0, weight=1)
        canvas_wrap.columnconfigure(0, weight=1)

        canvas = tk.Canvas(canvas_wrap, bg='#0d1117', highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_wrap, orient='vertical', command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg='#0d1117')
        scroll_frame.bind('<Configure>',
                          lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')

        def buscar_reportes(fecha_str):
            for w in scroll_frame.winfo_children():
                w.destroy()
            lbl_fecha_sel.config(text=f"Reportes del  {fecha_str}")

            archivo = "registro_reportes.xlsx"
            if not os.path.exists(archivo):
                ttk.Label(scroll_frame, text="No hay reportes guardados aún.",
                          foreground='#8b949e', font=('Helvetica', 10)
                          ).pack(pady=20)
                return
            try:
                df = pd.read_excel(archivo)
                df['_fecha'] = df['Fecha_Registro'].astype(str).str[:10]
                filtrado = df[df['_fecha'] == fecha_str]
            except Exception as ex:
                ttk.Label(scroll_frame, text=f"Error leyendo archivo: {ex}",
                          foreground='#e74c3c', font=('Helvetica', 9)
                          ).pack(pady=10)
                return

            if filtrado.empty:
                ttk.Label(scroll_frame,
                          text=f"No hay reportes para el {fecha_str}.",
                          foreground='#8b949e', font=('Helvetica', 10)
                          ).pack(pady=20)
                return

            ttk.Label(scroll_frame,
                      text=f"{len(filtrado)} reporte(s) — clic para ver detalle",
                      foreground='#8b949e', font=('Helvetica', 9)
                      ).pack(anchor='w', pady=(0, 8))

            for _, rd in filtrado.iterrows():
                def s(col, default=""):
                    v = rd.get(col, default)
                    return "" if pd.isna(v) else str(v)
                def n(col):
                    v = rd.get(col, 0)
                    return 0.0 if pd.isna(v) else float(v)

                turno_r = s("Turno", "—")
                hora_r  = s("Fecha_Registro")[11:16]
                fmt_r   = s("Formato_F1", "—")
                gly_r   = f"{n('GLY')}%"
                merma_r = f"{n('Merma')}%"
                hl_r    = s("HL Envasados", "—")

                card = tk.Frame(scroll_frame, bg='#161b22',
                                highlightbackground='#30363d', highlightthickness=1,
                                cursor='hand2')
                card.pack(fill='x', pady=4, ipadx=8, ipady=6)

                top = tk.Frame(card, bg='#161b22')
                top.pack(fill='x', padx=8, pady=(4, 2))
                tk.Label(top, text=f"Turno {turno_r}",
                         font=('Helvetica', 11, 'bold'),
                         fg='#C8102E', bg='#161b22').pack(side='left')
                tk.Label(top, text=f"  {hora_r}  ·  {fmt_r}",
                         font=('Helvetica', 9), fg='#8b949e', bg='#161b22').pack(side='left')

                kpi_row_f = tk.Frame(card, bg='#161b22')
                kpi_row_f.pack(fill='x', padx=8, pady=(2, 6))
                for lbl, val, kpi_key in [("HL Env.", hl_r, "HL Envasados"),
                                           ("GLY",    gly_r,   "GLY"),
                                           ("Merma",  merma_r, "Merma")]:
                    cel = tk.Frame(kpi_row_f, bg='#1c2333', padx=10, pady=4)
                    cel.pack(side='left', padx=(0, 6))
                    tk.Label(cel, text=lbl, font=('Helvetica', 8),
                             fg='#8b949e', bg='#1c2333').pack()
                    tk.Label(cel, text=val, font=('Helvetica', 12, 'bold'),
                             fg=get_color(kpi_key, val), bg='#1c2333').pack()

                def abrir_detalle(rd=rd):
                    res = {
                        "Cumplimiento": f"{n('Cumplimiento')}%",
                        "HL Envasados": s("HL Envasados"),
                        "GLY":          f"{n('GLY')}%",
                        "LEF":          f"{n('LEF')}%",
                        "Merma":        f"{n('Merma')}%",
                        "OSE":          f"{n('OSE')}%",
                        "Agua":         s("Agua"),
                        "CO2":          s("CO2"),
                        "Vapor":        s("Vapor"),
                        "TE (EPT)":     s("TE (EPT)"),
                    }
                    d_f1 = {
                        "Formato":        s("F1_Formato"),
                        "HL Programados": s("F1_HL_Programados"),
                        "HL Envasados":   s("F1_HL_Envasados"),
                        "Vol. Másico":    s("F1_Vol._Másico"),
                        "Merma":          s("F1_Merma"),
                    }
                    d_f2 = None
                    if s("F2_Formato"):
                        d_f2 = {
                            "Formato":        s("F2_Formato"),
                            "HL Programados": s("F2_HL_Programados"),
                            "HL Envasados":   s("F2_HL_Envasados"),
                            "Vol. Másico":    s("F2_Vol._Másico"),
                            "Merma":          s("F2_Merma"),
                        }
                    coms = {
                        "fallas":      s("Comentario_Fallas"),
                        "correctivas": s("Comentario_Correctivas"),
                        "cinco_w":     s("Comentario_5W"),
                        "avisos":      [],
                    }
                    mostrar_resultado(res, d_f1, d_f2, coms,
                                      datos_entrada=None,
                                      turno_label=s("Turno"),
                                      desde_historial=True)

                for widget in [card] + card.winfo_children():
                    widget.bind('<Button-1>', lambda e, f=abrir_detalle: f())

        # Buscar automáticamente al seleccionar fecha en el calendario
        cal.bind('<<CalendarSelected>>', lambda e: buscar_reportes(cal.get_date()))


if __name__ == "__main__":
    root = tk.Tk()
    create_ui(root, {})
    root.mainloop()