import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import os
import threading
import time
import concurrent.futures
from seniatOCR import SeniatOCR

class ExcelUploaderApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de RIFs SENIAT")
        self.geometry("500x600")
        self.config(bg="#f8f9fa")
        
        self.df_filtrado_global = None
        self.ruta_archivo_actual = None
        self.route_vars = {}
        
        self.chk_activo_var = tk.BooleanVar(value=True)
        self.chk_inactivo_var = tk.BooleanVar(value=False)
        
        self.label = tk.Label(
            self, 
            text="📁 Arrastra tu archivo Excel aquí\no\nusa el botón para buscarlo", 
            bg="#f8f9fa", 
            fg="#495057",
            font=("Arial", 11),
            pady=10
        )
        self.label.pack()
        
        self.btn_browse = tk.Button(
            self, 
            text="Buscar en archivos", 
            command=self.browse_file,
            font=("Arial", 10, "bold"),
            bg="#007bff",
            fg="white",
            padx=12,
            pady=6,
            relief="flat"
        )
        self.btn_browse.pack(pady=5)
        
        self.frame_checks = tk.Frame(self, bg="#f8f9fa")
        self.frame_checks.pack(pady=5, fill="both", expand=True)
        
        self.btn_download = tk.Button(
            self, 
            text="Iniciar Consulta", 
            command=self.iniciar_hilo_consultas,   
            font=("Arial", 10, "bold"),
            bg="#28a745",
            fg="white",
            padx=12,
            pady=6,
            relief="flat"
        )
        
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.handle_drop)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.process_excel(file_path)

    def handle_drop(self, event):
        file_path = event.data.strip('{}')
        if file_path.lower().endswith(('.xlsx')):
            self.process_excel(file_path)
        else:
            messagebox.showerror("Error", "Por favor, selecciona un archivo Excel válido (.xlsx)")

    def process_excel(self, file_path):
        try:
            self.ruta_archivo_actual = file_path
            
            df = pd.read_excel(file_path, header=8)
            nombres_buscados = tuple(f"T0X{i:02d}" for i in range(1, 19))
            df['Ruta (Nombre)'] = df['Ruta (Nombre)'].astype(str)
            
            self.df_filtrado_global = df[df['Ruta (Nombre)'].str.startswith(nombres_buscados)].copy()
            
            self.df_filtrado_global['Cliente (Identificación)'] = (
                self.df_filtrado_global['Cliente (Identificación)']
                .astype(str)
                .str.replace('-', '', regex=False)
                .str.strip()
            )
            
            for widget in self.frame_checks.winfo_children():
                widget.destroy()
            self.route_vars.clear()

            rutas_counts = self.df_filtrado_global['Ruta (Nombre)'].value_counts().to_dict()
            rutas_encontradas = sorted(rutas_counts.keys())

            if not rutas_encontradas:
                messagebox.showwarning("Aviso", "No se encontraron rutas coincidentes.")
                self.btn_download.pack_forget()
                return

            lbl_estatus = tk.Label(self.frame_checks, text="Selecciona el/los estatus a evaluar:", bg="#f8f9fa", font=("Arial", 9, "bold"), fg="#333")
            lbl_estatus.pack(anchor="w", padx=20, pady=2)
            
            frame_estatus = tk.Frame(self.frame_checks, bg="#f8f9fa")
            frame_estatus.pack(anchor="w", padx=30, pady=2)
            
            tk.Checkbutton(frame_estatus, text="Activos", variable=self.chk_activo_var, bg="#f8f9fa").pack(side="left", padx=5)
            tk.Checkbutton(frame_estatus, text="Inactivos", variable=self.chk_inactivo_var, bg="#f8f9fa").pack(side="left", padx=5)

            lbl_sub = tk.Label(self.frame_checks, text="Selecciona las rutas a incluir:", bg="#f8f9fa", font=("Arial", 9, "bold"), fg="#333")
            lbl_sub.pack(anchor="w", padx=20, pady=(8, 2))

            container_scroll = tk.Frame(self.frame_checks, bg="#f8f9fa", bd=1, relief="solid")
            container_scroll.pack(fill="both", expand=True, padx=25, pady=5)

            canvas = tk.Canvas(container_scroll, bg="white", highlightthickness=0)
            scrollbar = tk.Scrollbar(container_scroll, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg="white")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            for ruta in rutas_encontradas:
                var = tk.BooleanVar(value=True)
                self.route_vars[ruta] = var
                texto_ruta = f"{ruta} ({rutas_counts[ruta]} clientes)"
                chk = tk.Checkbutton(
                    scrollable_frame, 
                    text=texto_ruta, 
                    variable=var, 
                    bg="white",
                    font=("Arial", 9)
                )
                chk.pack(anchor="w", padx=10, pady=2)

            if not self.btn_download.winfo_ismapped():
                self.btn_download.pack(pady=10)
            
            messagebox.showinfo("Éxito", f"Archivo cargado. {len(rutas_encontradas)} rutas listas.")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def escribir_log(self, mensaje):
        self.after(0, self._insertar_texto, mensaje)

    def _insertar_texto(self, mensaje):
        self.txt_log.config(state="normal")
        self.txt_log.insert(tk.END, mensaje + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def iniciar_hilo_consultas(self):
        rutas_seleccionadas = [ruta for ruta, var in self.route_vars.items() if var.get()]
        if not rutas_seleccionadas:
            messagebox.showwarning("Atención", "Debes seleccionar al menos una ruta.")
            return

        estatus_permitidos = []
        if self.chk_activo_var.get(): estatus_permitidos.append("activo")
        if self.chk_inactivo_var.get(): estatus_permitidos.append("inactivo")
        if not estatus_permitidos:
            messagebox.showwarning("Atención", "Debes seleccionar al menos un estatus.")
            return

        df_final = self.df_filtrado_global[self.df_filtrado_global['Ruta (Nombre)'].isin(rutas_seleccionadas)]
        if 'Estatus' in df_final.columns:
            df_final = df_final[df_final['Estatus'].astype(str).str.strip().str.lower().isin(estatus_permitidos)]

        rifs_a_consultar = df_final['Cliente (Identificación)'].dropna().unique().tolist() 
        
        if not rifs_a_consultar:
            messagebox.showwarning("Aviso", "No hay RIFs para consultar.")
            return

        self.top_log = tk.Toplevel(self)
        self.top_log.title("Consola SENIAT - Procesamiento por Lotes")
        self.top_log.geometry("650x450")
        self.top_log.config(bg="#1e1e1e")
        
        self.txt_log = tk.Text(self.top_log, bg="#1e1e1e", fg="#4af626", font=("Consolas", 10))
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=10)
        self.txt_log.config(state="disabled")
        
        self.btn_download.config(state="disabled")

        hilo_maestro = threading.Thread(
            target=self.gestor_multihilo_background, 
            args=(rifs_a_consultar,), 
            daemon=True
        )
        hilo_maestro.start()

    def procesar_lote_hilo(self, id_hilo, chunk_rifs, mapa_contribuyentes, lock, contador_progreso, total_rifs):
        api = SeniatOCR() 
        self.escribir_log(f"[Hilo #{id_hilo}] Iniciado. Lote asignado: {len(chunk_rifs)} RIFs.")

        for rif in chunk_rifs:
            max_intentos = 12
            intento = 1
            res = None
            
            while intento <= max_intentos:
                res = api.consultar_rif(rif) 
                
                if res and res.get("status") in ["success", "not_found"]:
                    if res.get("status") == "success":
                        self.escribir_log(f"[{rif}] ✅ Captcha Correcto.")
                    break
                    
                if res and res.get("status") == "fail_ocr":
                    self.escribir_log(f"[{rif}] ⚠️ Captcha Incorrecto (Intento {intento}/{max_intentos}). Reintentando...")
                    intento += 1
                    time.sleep(1.0)
                    continue
                    
                self.escribir_log(f"[{rif}] ❌ Error de conexión o servidor.")
                break

            es_contribuyente = "NO"
            if res and res.get("status") == "success":
                nombre = res.get("nombre", "SIN NOMBRE").title()
                condicion = res.get("condicion", "").upper()
                
                if "AGENTE DE RETENCIÓN DEL IVA" in condicion:
                    self.escribir_log(f"[{rif}] ▶ Procesado: {nombre} | Especial: SI")
                    es_contribuyente = "SI"
                else:
                    self.escribir_log(f"[{rif}] ▶ Procesado: {nombre} | Especial: NO")
            else:
                self.escribir_log(f"[{rif}] ▶ El RIF No Existe o fallaron los intentos.")
                
            with lock:
                mapa_contribuyentes[rif] = es_contribuyente
                contador_progreso[0] += 1
                completados = contador_progreso[0]
                
            if completados % 10 == 0 or completados == total_rifs:
                self.escribir_log(f"\n--- PROGRESO GENERAL: {completados} / {total_rifs} COMPLETADOS ---\n")
                
            time.sleep(0.5)

    def gestor_multihilo_background(self, rifs_a_consultar):
        total = len(rifs_a_consultar)
        MAX_WORKERS = 8
        
        self.escribir_log(f"--- INICIANDO DISTRIBUCIÓN POR LOTES ({total} REGISTROS) ---")
        self.escribir_log(f"--- HILOS ACTIVOS: {MAX_WORKERS} ---\n")
        self.escribir_log(f"--- Inicializando motor de reconocimiento visual... ---\n")
        
        # Dividir la lista de RIFs en 'MAX_WORKERS' partes iguales pue
        k, m = divmod(total, MAX_WORKERS)
        chunks = [rifs_a_consultar[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(MAX_WORKERS)]
        # Filtrar chunks vacíos por si hay menos RIFs q hilos
        chunks = [c for c in chunks if len(c) > 0]
        
        mapa_contribuyentes = {}
        lock = threading.Lock()
        contador_progreso = [0]
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(chunks)) as executor:
            futuros = []
            for idx, chunk in enumerate(chunks, 1):
                futuro = executor.submit(
                    self.procesar_lote_hilo, 
                    idx, 
                    chunk, 
                    mapa_contribuyentes, 
                    lock, 
                    contador_progreso, 
                    total
                )
                futuros.append(futuro)
                
            concurrent.futures.wait(futuros)

        self.generar_excel_final(mapa_contribuyentes)

    def generar_excel_final(self, mapa_contribuyentes):
        self.escribir_log("\n--- CONSULTAS FINALIZADAS ---")
        self.escribir_log("Inyectando resultados en el documento Excel...")
        try:
            wb = openpyxl.load_workbook(self.ruta_archivo_actual)
            ws = wb.active
            
            header_row = 9
            col_rif = None
            
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=header_row, column=col).value == "Cliente (Identificación)":
                    col_rif = col
                    break
            
            if col_rif is not None:
                new_col = ws.max_column + 1
                letra_nueva_col = get_column_letter(new_col)
                
                celda_header = ws.cell(row=header_row, column=new_col, value="Contribuyente")
                celda_referencia = ws.cell(row=header_row, column=new_col - 1)
                
                if celda_referencia.has_style:
                    celda_header.font = copy(celda_referencia.font)
                    celda_header.border = copy(celda_referencia.border)
                    celda_header.fill = copy(celda_referencia.fill)
                    celda_header.alignment = copy(celda_referencia.alignment)
                
                ws.column_dimensions[letra_nueva_col].width = 16
                
                for row in range(header_row + 1, ws.max_row + 1):
                    valor_rif = str(ws.cell(row=row, column=col_rif).value or "").replace("-", "").strip()
                    if valor_rif and valor_rif in mapa_contribuyentes:
                        celda_resultado = ws.cell(row=row, column=new_col, value=mapa_contribuyentes[valor_rif])
                        celda_resultado.alignment = openpyxl.styles.Alignment(horizontal='center')
                
                base, ext = os.path.splitext(self.ruta_archivo_actual)
                nuevo_archivo = f"{base}_PROCESADO{ext}"
                wb.save(nuevo_archivo)
                
                self.escribir_log(f"¡ÉXITO! Archivo generado correctamente:\n{nuevo_archivo}")
                self.after(0, lambda: messagebox.showinfo("Completado", f"Archivo procesado guardado:\n\n{nuevo_archivo}"))
            else:
                self.escribir_log("Error: No se encontró la columna de identificación.")
                
        except Exception as e:
            self.escribir_log(f"Error fatal guardando Excel: {e}")
        finally:
            self.after(0, lambda: self.btn_download.config(state="normal"))

if __name__ == "__main__":
    app = ExcelUploaderApp()
    app.mainloop()