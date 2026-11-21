import pandas as pd
from tkinter import Tk, filedialog

root = Tk()
root.withdraw()

archivo = filedialog.askopenfilename(
    title="Seleccione un archivo Excel",
    filetypes=[
        ("Excel files", "*.xlsx *.xls"),
        ("All files", "*.*")
    ]
)

if not archivo:
    print("No se seleccionó ningún archivo.")
    exit()

df = pd.read_excel(archivo)

print(f"Archivo seleccionado: {archivo}")
print(df.head())