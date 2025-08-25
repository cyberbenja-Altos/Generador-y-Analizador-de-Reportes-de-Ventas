import os.path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

from main import generar_excel_completo


def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="selecciona el archivo CSV(excel)",
        filetypes =[("archivos CSV", "*.csv")]
    )
    if ruta:
        entry_cvs.delete(0, tk.END)
        entry_cvs.insert(0, ruta)

def generar_reporte():
    archivo_cvs = entry_cvs.get()
    if not archivo_cvs or not os.path.exists(archivo_cvs):
        messagebox.showerror("error", "selecciona un archivo CSV valido.")
        return

    try:
        df = pd.read_csv(archivo_cvs)
        df['fecha'] = pd.to_datetime(df['fecha'])

        nombre_excel = os.path.splitext(archivo_cvs)[0] + "_reporte.xlsx"
        generar_excel_completo(df, nombre_excel)

        messagebox.showinfo("Exito", f"reporte generado:\n{nombre_excel}")
    except Exception as e:
        messagebox.showerror("Error, " f"Ocurrio un problema:\n{e}")

    #interfaz
root = tk.Tk()
root.title("generador de reportes de ventas")
root.geometry("900x400")

tk.Label(root, text="archivo CVS:").pack(pady=5)
entry_cvs = tk.Entry(root, width=50)
entry_cvs.pack(pady=5)

tk.Button(root,  text="seleccionar CVS,", command=seleccionar_archivo).pack(pady=5)

tk.Button(root, text="generar reporte", command=generar_reporte, bg='green', fg="white").pack(pady=10)

root.mainloop()