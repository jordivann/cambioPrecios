import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import mysql.connector
from datetime import datetime

# Configuración de la base de datos
DB_CONFIG = {
    "host": "serverndv.no-ip.org",
    "port": 6613,
    "user": "comprasfsa",
    "password": "S1nch2z@c4mpr1sfs1",
    "database": "plexdr",
    "charset": "utf8"
}

class ProductReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Informe de Productos Actualizados")
        self.root.geometry("500x300")
        self.create_widgets()

    def create_widgets(self):
        # Selección de fechas
        frame = ttk.LabelFrame(self.root, text="Seleccione el rango de fechas", padding=(20, 10))
        frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame, text="Fecha desde:").grid(column=0, row=0, padx=5, pady=5, sticky="w")
        self.start_date = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern="y-mm-dd")
        self.start_date.grid(column=1, row=0, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="Fecha hasta:").grid(column=0, row=1, padx=5, pady=5, sticky="w")
        self.end_date = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern="y-mm-dd")
        self.end_date.grid(column=1, row=1, padx=5, pady=5, sticky="w")

        # Botón para generar el informe
        generate_btn = ttk.Button(self.root, text="Generar Informe", command=self.generate_report)
        generate_btn.pack(pady=20)

    def generate_report(self):
        # Obtener las fechas seleccionadas
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()

        if start_date > end_date:
            messagebox.showerror("Error", "La fecha de inicio no puede ser posterior a la fecha final.")
            return

        try:
            # Conectar a la base de datos
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()

            # Consulta para obtener los datos
            query = """
            SELECT IDProducto, Producto, Presentacion, FechaModificacion, UltimoPrecio, MargenPVP
            FROM productos
            WHERE FechaModificacion BETWEEN %s AND %s
            """
            cursor.execute(query, (start_date, end_date))
            data = cursor.fetchall()

            # Verificar si hay resultados
            if not data:
                messagebox.showinfo("Sin resultados", "No se encontraron productos actualizados en el rango de fechas seleccionado.")
                return

            # Convertir los datos a un DataFrame de pandas
            columns = ["IDProducto", "Producto", "Presentacion", "FechaModificacion", "UltimoPrecio", "MargenPVP"]
            df = pd.DataFrame(data, columns=columns)

            # Guardar el informe en un archivo Excel
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Éxito", f"Informe guardado en {save_path}")

        except mysql.connector.Error as e:
            messagebox.showerror("Error de conexión", f"No se pudo conectar a la base de datos:\n{e}")
        finally:
            cursor.close()
            conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductReportApp(root)
    root.mainloop()
