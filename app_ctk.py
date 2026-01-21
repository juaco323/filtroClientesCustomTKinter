import customtkinter as ctk
import pandas as pd
import sqlite3
import os
import webbrowser
import ipaddress
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter as tk

# Configuracion de apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DB_FILE = "sistema_isp_db"

class BuscadorDRECH(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Sistema de Gestion DRECH")
        self.geometry("1200x700")
        self.minsize(900, 500)
        
        # Frame principal
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Sidebar
        self.crear_sidebar()
        
        # Panel principal
        self.crear_panel_principal()
        
        # Cargar datos iniciales
        self.actualizar_contador()
        self.cargar_ubicaciones()  # Cargar ubicaciones si ya existe BD
    
    def crear_sidebar(self):
        # Frame lateral
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)
        
        # Titulo sidebar
        self.logo_label = ctk.CTkLabel(self.sidebar, text="Carga de Datos", 
                                        font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Info
        self.info_label = ctk.CTkLabel(self.sidebar, text="Sube aqui el archivo\nConsolidado.xlsx",
                                        font=ctk.CTkFont(size=12))
        self.info_label.grid(row=1, column=0, padx=20, pady=10)
        
        # Boton seleccionar archivo
        self.btn_seleccionar = ctk.CTkButton(self.sidebar, text="Seleccionar Excel",
                                              command=self.seleccionar_archivo)
        self.btn_seleccionar.grid(row=2, column=0, padx=20, pady=10)
        
        # Label archivo seleccionado
        self.archivo_label = ctk.CTkLabel(self.sidebar, text="Ningun archivo seleccionado",
                                           font=ctk.CTkFont(size=11), wraplength=200)
        self.archivo_label.grid(row=3, column=0, padx=20, pady=5)
        
        # Boton procesar
        self.btn_procesar = ctk.CTkButton(self.sidebar, text="Procesar y Actualizar BD",
                                           command=self.procesar_archivo, state="disabled",
                                           fg_color="#28a745", hover_color="#218838")
        self.btn_procesar.grid(row=4, column=0, padx=20, pady=10)
        
        # Boton limpiar BD
        self.btn_limpiar = ctk.CTkButton(self.sidebar, text="Limpiar BD",
                                          command=self.limpiar_bd,
                                          fg_color="#dc3545", hover_color="#c82333",
                                          width=140)
        self.btn_limpiar.grid(row=5, column=0, padx=20, pady=5)
        
        # Separador
        self.separator = ctk.CTkFrame(self.sidebar, height=2, fg_color="gray50")
        self.separator.grid(row=6, column=0, padx=20, pady=20, sticky="ew")
        
        # Contador clientes
        self.clientes_label = ctk.CTkLabel(self.sidebar, text="Clientes Totales",
                                            font=ctk.CTkFont(size=14))
        self.clientes_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        
        self.contador_label = ctk.CTkLabel(self.sidebar, text="0",
                                            font=ctk.CTkFont(size=36, weight="bold"))
        self.contador_label.grid(row=8, column=0, padx=20, pady=(0, 20))
        
        self.archivo_seleccionado = None
    
    def crear_panel_principal(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)
        
        # Titulo
        self.titulo = ctk.CTkLabel(self.main_frame, text="Sistema de Gestion DRECH",
                                    font=ctk.CTkFont(size=28, weight="bold"))
        self.titulo.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # Frame busqueda
        self.search_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.search_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        self.search_frame.grid_columnconfigure(0, weight=1)
        
        # Label busqueda
        self.search_label = ctk.CTkLabel(self.search_frame, text="Busque por Nombre del Cliente o IP",
                                          font=ctk.CTkFont(size=14))
        self.search_label.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        # Frame para input y boton
        self.input_frame = ctk.CTkFrame(self.search_frame, fg_color="transparent")
        self.input_frame.grid(row=1, column=0, sticky="ew")
        self.input_frame.grid_columnconfigure(0, weight=1)
        
        # Input busqueda
        self.search_entry = ctk.CTkEntry(self.input_frame, placeholder_text="Ingresa nombre o IP",
                                          height=40, font=ctk.CTkFont(size=14))
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.search_entry.bind("<Return>", lambda e: self.buscar())
        
        # Boton buscar
        self.btn_buscar = ctk.CTkButton(self.input_frame, text="🔍 Buscar", width=120,
                                         height=40, command=self.buscar)
        self.btn_buscar.grid(row=0, column=1)
        
        # Frame para filtros (inicialmente oculto)
        self.filter_frame = ctk.CTkFrame(self.search_frame, fg_color="transparent")
        self.filter_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        # Botón para mostrar/ocultar filtros
        self.btn_filtrar = ctk.CTkButton(self.filter_frame, text="🔽 Filtrar", width=100,
                                          height=30, command=self.toggle_filtros,
                                          fg_color="#555555", hover_color="#666666")
        self.btn_filtrar.grid(row=0, column=0, padx=(0, 10))
        
        # Frame contenedor de filtros (inicialmente oculto)
        self.filtros_container = ctk.CTkFrame(self.filter_frame, fg_color="transparent")
        self.filtros_ocultos = True  # Estado inicial: ocultos
        
        # Label filtro ubicación
        self.filter_ubicacion_label = ctk.CTkLabel(self.filtros_container, text="Ubicación:",
                                          font=ctk.CTkFont(size=13))
        self.filter_ubicacion_label.grid(row=0, column=0, padx=(0, 5))
        
        # ComboBox de ubicaciones
        self.ubicacion_var = ctk.StringVar(value="Todas las ubicaciones")
        self.ubicacion_combo = ctk.CTkComboBox(self.filtros_container, 
                                                 variable=self.ubicacion_var,
                                                 values=["Todas las ubicaciones"],
                                                 width=200,
                                                 height=35,
                                                 font=ctk.CTkFont(size=12),
                                                 state="readonly")
        self.ubicacion_combo.grid(row=0, column=1, padx=(0, 15))
        
        # Label filtro zona
        self.filter_zona_label = ctk.CTkLabel(self.filtros_container, text="Zona:",
                                          font=ctk.CTkFont(size=13))
        self.filter_zona_label.grid(row=0, column=2, padx=(0, 5))
        
        # ComboBox de zonas
        self.zona_var = ctk.StringVar(value="Todas las zonas")
        self.zona_combo = ctk.CTkComboBox(self.filtros_container, 
                                           variable=self.zona_var,
                                           values=["Todas las zonas"],
                                           width=200,
                                           height=35,
                                           font=ctk.CTkFont(size=12),
                                           state="readonly")
        self.zona_combo.grid(row=0, column=3, padx=(0, 15))
        
        # Botón limpiar filtros
        self.btn_limpiar_filtros = ctk.CTkButton(self.filtros_container, text="✖ Limpiar", 
                                                  width=80, height=30,
                                                  command=self.limpiar_filtros,
                                                  fg_color="#dc3545", hover_color="#c82333")
        self.btn_limpiar_filtros.grid(row=0, column=4)
        
        # Frame para tabla con scrollbar
        self.table_frame = ctk.CTkFrame(self.main_frame)
        self.table_frame.grid(row=2, column=0, sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)
        
        # Crear Treeview (tabla)
        self.crear_tabla()
    
    def crear_tabla(self):
        # Estilo de la tabla
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#2b2b2b",
                        foreground="white",
                        rowheight=30,
                        fieldbackground="#2b2b2b",
                        font=('Segoe UI', 11))
        style.configure("Treeview.Heading",
                        background="#1f538d",
                        foreground="white",
                        font=('Segoe UI', 11, 'bold'))
        style.map("Treeview",
                  background=[('selected', '#1f538d')])
        
        # Columnas
        columnas = ("cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro")
        
        self.tabla = ttk.Treeview(self.table_frame, columns=columnas, show="headings", selectmode="browse")
        
        # Tag para hipervinculos (azul)
        self.tabla.tag_configure("link", foreground="#4da6ff")
        
        # Configurar columnas
        self.tabla.heading("cliente", text="Nombre Cliente")
        self.tabla.heading("ip_antena", text="IP Antena 🔗")
        self.tabla.heading("ip_router", text="IP Router 🔗")
        self.tabla.heading("ubicacion", text="Ubicacion")
        self.tabla.heading("zona", text="Zona")
        self.tabla.heading("plan", text="Plan")
        self.tabla.heading("fecha_registro", text="Fecha Reg.")
        
        # Ancho de columnas
        self.tabla.column("cliente", width=150, minwidth=100)
        self.tabla.column("ip_antena", width=100, minwidth=80)
        self.tabla.column("ip_router", width=100, minwidth=80)
        self.tabla.column("ubicacion", width=120, minwidth=80)
        self.tabla.column("zona", width=100, minwidth=80)
        self.tabla.column("plan", width=80, minwidth=60)
        self.tabla.column("fecha_registro", width=100, minwidth=80)
        
        # Scrollbars
        scrollbar_y = ctk.CTkScrollbar(self.table_frame, command=self.tabla.yview)
        scrollbar_x = ctk.CTkScrollbar(self.table_frame, command=self.tabla.xview, orientation="horizontal")
        self.tabla.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Grid
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # Bind doble click para abrir IP
        self.tabla.bind("<Double-1>", self.abrir_ip)
    
    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo_seleccionado = archivo
            nombre = os.path.basename(archivo)
            self.archivo_label.configure(text=f"📄 {nombre}")
            self.btn_procesar.configure(state="normal")
    
    def procesar_archivo(self):
        if not self.archivo_seleccionado:
            return
        
        try:
            exito, mensaje = self.actualizar_db(self.archivo_seleccionado)
            if exito:
                messagebox.showinfo("Exito", mensaje)
                self.actualizar_contador()
                self.cargar_ubicaciones()  # Cargar ubicaciones automáticamente
            else:
                messagebox.showerror("Error", mensaje)
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar: {e}")
    
    def actualizar_db(self, excel_path):
        try:
            xlsx = pd.ExcelFile(excel_path)
            dfs = []
            
            for sheet_name in xlsx.sheet_names:
                df_sheet = pd.read_excel(xlsx, sheet_name=sheet_name, dtype=str)
                if not df_sheet.empty:
                    # Normalizar columnas
                    df_sheet.columns = (df_sheet.columns
                          .str.strip()
                          .str.lower()
                          .str.replace(' ', '_', regex=False)
                          .str.replace('ó', 'o', regex=False)
                          .str.replace('í', 'i', regex=False)
                          .str.replace('á', 'a', regex=False)
                          .str.replace('é', 'e', regex=False)
                          .str.replace('ú', 'u', regex=False)
                          .str.replace('.', '', regex=False)
                          .str.replace('º', '', regex=False)
                          .str.replace('n°', 'n', regex=False))
                    
                    if 'zona' not in df_sheet.columns:
                        df_sheet['zona'] = sheet_name
                    dfs.append(df_sheet)
            
            if not dfs:
                return False, "El archivo Excel esta vacio."
            
            df = pd.concat(dfs, ignore_index=True)
            
            if 'cliente' in df.columns:
                df = df.dropna(subset=['cliente'])
                df = df[df['cliente'].str.strip() != '']
            
            if 'fecha_registro' in df.columns:
                df['fecha_registro'] = pd.to_datetime(df['fecha_registro'], errors='coerce').dt.strftime('%d-%m-%Y')
            
            conn = sqlite3.connect(DB_FILE)
            df.to_sql('clientes', conn, if_exists='replace', index=False)
            
            cursor = conn.cursor()
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_cliente ON clientes (cliente)")
            conn.commit()
            conn.close()
            
            return True, f"Base de datos actualizada con {len(df)} registros."
        except Exception as e:
            return False, f"Error al actualizar BD: {e}"
    
    def limpiar_bd(self):
        if not os.path.exists(DB_FILE):
            messagebox.showinfo("Info", "No hay base de datos para eliminar.")
            return
        
        confirmar = messagebox.askyesno("Confirmar", "¿Estas seguro de eliminar toda la base de datos?\nEsta accion no se puede deshacer.")
        if confirmar:
            try:
                os.remove(DB_FILE)
                # Limpiar tabla
                for item in self.tabla.get_children():
                    self.tabla.delete(item)
                self.actualizar_contador()
                messagebox.showinfo("Exito", "Base de datos eliminada correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo eliminar: {e}")
    
    def actualizar_contador(self):
        if os.path.exists(DB_FILE):
            try:
                conn = sqlite3.connect(DB_FILE)
                total = pd.read_sql_query("SELECT COUNT(*) as total FROM clientes", conn)['total'][0]
                self.contador_label.configure(text=str(total))
                conn.close()
            except:
                self.contador_label.configure(text="0")
        else:
            self.contador_label.configure(text="0")
    
    def cargar_ubicaciones(self):
        """Carga las ubicaciones y zonas únicas desde la base de datos a los ComboBox"""
        if not os.path.exists(DB_FILE):
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            
            # Verificar columnas disponibles
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(clientes)")
            columnas = [col[1] for col in cursor.fetchall()]
            
            # Cargar ubicaciones
            if 'ubicacion' in columnas:
                df_ubicaciones = pd.read_sql_query(
                    "SELECT DISTINCT TRIM(ubicacion) as ubicacion FROM clientes WHERE ubicacion IS NOT NULL AND TRIM(ubicacion) != '' ORDER BY ubicacion", 
                    conn
                )
                ubicaciones = ["Todas las ubicaciones"] + df_ubicaciones['ubicacion'].tolist()
                self.ubicacion_combo.configure(values=ubicaciones)
                self.ubicacion_var.set("Todas las ubicaciones")
            
            # Cargar zonas
            if 'zona' in columnas:
                df_zonas = pd.read_sql_query(
                    "SELECT DISTINCT TRIM(zona) as zona FROM clientes WHERE zona IS NOT NULL AND TRIM(zona) != '' ORDER BY zona", 
                    conn
                )
                zonas = ["Todas las zonas"] + df_zonas['zona'].tolist()
                self.zona_combo.configure(values=zonas)
                self.zona_var.set("Todas las zonas")
            
            conn.close()
            
        except Exception as e:
            print(f"Error al cargar ubicaciones/zonas: {e}")
    
    def toggle_filtros(self):
        """Muestra u oculta el panel de filtros"""
        if self.filtros_ocultos:
            self.filtros_container.grid(row=0, column=1, padx=(10, 0))
            self.btn_filtrar.configure(text="🔼 Filtrar")
            self.filtros_ocultos = False
        else:
            self.filtros_container.grid_forget()
            self.btn_filtrar.configure(text="🔽 Filtrar")
            self.filtros_ocultos = True
    
    def limpiar_filtros(self):
        """Limpia todos los filtros seleccionados"""
        self.ubicacion_var.set("Todas las ubicaciones")
        self.zona_var.set("Todas las zonas")
    
    def buscar(self):
        query = self.search_entry.get().strip()
        ubicacion_filtro = self.ubicacion_var.get()
        zona_filtro = self.zona_var.get()
        
        # Limpiar tabla
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        
        if not query:
            messagebox.showwarning("Aviso", "Ingresa un nombre o IP para buscar.")
            return
        
        if not os.path.exists(DB_FILE):
            messagebox.showwarning("Aviso", "No hay base de datos. Carga un archivo Excel primero.")
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            
            # Obtener columnas disponibles
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(clientes)")
            columnas_disponibles = [col[1] for col in cursor.fetchall()]
            
            columnas_deseadas = ['cliente', 'ip_antena', 'ip_router', 'ubicacion', 'plan', 'fecha_registro', 'zona']
            columnas_select = [col for col in columnas_deseadas if col in columnas_disponibles]
            
            # Construir consulta SQL con filtros
            query_escaped = query.replace("'", "''")
            
            sql = f"""
            SELECT {', '.join(columnas_select)}
            FROM clientes 
            WHERE (LOWER(cliente) LIKE LOWER('%{query_escaped}%') 
               OR LOWER(ip_antena) LIKE LOWER('%{query_escaped}%')
               OR LOWER(ubicacion) LIKE LOWER('%{query_escaped}%'))
            """
            
            # Agregar filtro de ubicación si no es "Todas las ubicaciones"
            if ubicacion_filtro != "Todas las ubicaciones":
                ubicacion_escaped = ubicacion_filtro.strip().replace("'", "''")
                sql += f" AND LOWER(TRIM(ubicacion)) = LOWER(TRIM('{ubicacion_escaped}'))"
            
            # Agregar filtro de zona si no es "Todas las zonas"
            if zona_filtro != "Todas las zonas":
                zona_escaped = zona_filtro.strip().replace("'", "''")
                sql += f" AND LOWER(TRIM(zona)) = LOWER(TRIM('{zona_escaped}'))"
            
            df = pd.read_sql_query(sql, conn)
            conn.close()
            
            if df.empty:
                messagebox.showinfo("Resultado", "No se encontraron coincidencias.")
                return
            
            # Insertar datos en tabla con formato de links
            for _, row in df.iterrows():
                valores = []
                for col in ["cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro"]:
                    if col in df.columns:
                        val = row[col] if pd.notnull(row[col]) else ""
                        # Agregar indicador visual de link para IPs
                        if col in ["ip_antena", "ip_router"] and val and str(val).strip():
                            val = f"🔗 {val}"
                        valores.append(val)
                    else:
                        valores.append("")
                self.tabla.insert("", "end", values=valores, tags=("link",))
                
        except Exception as e:
            messagebox.showerror("Error", f"Error en la busqueda: {e}")
    
    def es_ip_valida(self, ip_str):
        """Valida si una cadena es una dirección IP válida"""
        try:
            ipaddress.ip_address(ip_str)
            return True
        except ValueError:
            return False
    
    def abrir_ip(self, event):
        item = self.tabla.selection()
        if not item:
            return
        
        # Obtener columna clickeada
        region = self.tabla.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        col = self.tabla.identify_column(event.x)
        col_index = int(col.replace("#", "")) - 1
        
        # Columnas de IP son indice 1 (ip_antena) y 2 (ip_router)
        if col_index in [1, 2]:
            valores = self.tabla.item(item[0], "values")
            ip = valores[col_index]
            if ip and str(ip).strip():
                # Remover el emoji del link si existe
                ip_limpia = str(ip).replace("🔗 ", "").strip()
                if ip_limpia:
                    # Validar con módulo ipaddress
                    if self.es_ip_valida(ip_limpia):
                        url = f"http://{ip_limpia}"
                        webbrowser.open(url)
                    else:
                        messagebox.showinfo("Info", f"'{ip_limpia}' no es una dirección IP válida.")

if __name__ == "__main__":
    app = BuscadorDRECH()
    app.mainloop()
