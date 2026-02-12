import customtkinter as ctk
import pandas as pd
import sqlite3
import os
import webbrowser
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter as tk

# Configuracion de apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Guardar la base de datos en AppData para que persista
def obtener_ruta_db():
    """Obtiene la ruta de la base de datos en AppData (persistente)"""
    appdata = os.getenv('APPDATA')
    if appdata:
        carpeta_app = os.path.join(appdata, 'BuscadorDRECH')
        if not os.path.exists(carpeta_app):
            os.makedirs(carpeta_app)
        return os.path.join(carpeta_app, 'sistema_isp_db')
    else:
        # Fallback: usar directorio actual
        return 'sistema_isp_db'

DB_FILE = obtener_ruta_db()

def formatear_fecha(event):
    """Formatear fecha automaticamente: usuario escribe numeros y se agregan guiones"""
    widget = event.widget
    texto = widget.get()
    
    # Remover todo excepto numeros
    solo_numeros = ''.join(c for c in texto if c.isdigit())
    
    # Limitar a 8 digitos (DDMMYYYY)
    solo_numeros = solo_numeros[:8]
    
    # Formatear con guiones
    if len(solo_numeros) <= 2:
        nuevo_texto = solo_numeros
    elif len(solo_numeros) <= 4:
        nuevo_texto = f"{solo_numeros[:2]}-{solo_numeros[2:]}"
    else:
        nuevo_texto = f"{solo_numeros[:2]}-{solo_numeros[2:4]}-{solo_numeros[4:]}"
    
    # Solo actualizar si cambio
    if texto != nuevo_texto:
        widget.delete(0, tk.END)
        widget.insert(0, nuevo_texto)
        # Mover cursor al final
        widget.icursor(tk.END)

class BuscadorDRECH(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Sistema de Gestion DRECH")
        self.geometry("1200x700")
        self.minsize(900, 500)
        
        # Frame principal
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Variable para modo edicion
        self.modo_admin = False
        self.cliente_editando = None
        
        # Sidebar
        self.crear_sidebar()
        
        # Panel principal
        self.crear_panel_principal()
        
        # Cargar datos iniciales
        self.actualizar_contador()
    
    def crear_sidebar(self):
        # Frame lateral
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(8, weight=1)  # Frame de zonas expandible
        
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
        self.btn_limpiar = ctk.CTkButton(self.sidebar, text="🗑️ Limpiar BD",
                                          command=self.limpiar_bd,
                                          fg_color="#dc3545", hover_color="#c82333",
                                          width=140)
        self.btn_limpiar.grid(row=5, column=0, padx=20, pady=5)
        
        # Separador
        self.separator = ctk.CTkFrame(self.sidebar, height=2, fg_color="gray50")
        self.separator.grid(row=6, column=0, padx=20, pady=15, sticky="ew")
        
        # Titulo Zonas
        self.zonas_titulo = ctk.CTkLabel(self.sidebar, text="📍 Zonas",
                                          font=ctk.CTkFont(size=16, weight="bold"))
        self.zonas_titulo.grid(row=7, column=0, padx=20, pady=(5, 10))
        
        # Frame scrollable para zonas
        self.zonas_frame = ctk.CTkScrollableFrame(self.sidebar, width=200, height=150,
                                                   fg_color="transparent")
        self.zonas_frame.grid(row=8, column=0, padx=10, pady=(0, 10), sticky="nsew")
        
        # Diccionario para guardar los botones de zonas
        self.zonas_buttons = {}
        
        # Separador 2
        self.separator2 = ctk.CTkFrame(self.sidebar, height=2, fg_color="gray50")
        self.separator2.grid(row=9, column=0, padx=20, pady=10, sticky="ew")
        
        # Contador clientes
        self.clientes_label = ctk.CTkLabel(self.sidebar, text="Clientes Totales",
                                            font=ctk.CTkFont(size=14))
        self.clientes_label.grid(row=10, column=0, padx=20, pady=(10, 0))
        
        self.contador_label = ctk.CTkLabel(self.sidebar, text="0",
                                            font=ctk.CTkFont(size=36, weight="bold"))
        self.contador_label.grid(row=11, column=0, padx=20, pady=(0, 10))
        
        # Boton Modificar datos BD (en la parte inferior)
        self.btn_admin = ctk.CTkButton(self.sidebar, text="📝 Modificar datos BD",
                                        command=self.toggle_modo_admin,
                                        fg_color="#6c757d", hover_color="#5a6268",
                                        height=40)
        self.btn_admin.grid(row=12, column=0, padx=20, pady=(10, 20), sticky="s")
        
        self.archivo_seleccionado = None
        self.zona_seleccionada = None
    
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
        
        # Mensaje de ayuda (oculto por defecto)
        self.help_message = ctk.CTkLabel(self.search_frame, text="",
                                          font=ctk.CTkFont(size=12), text_color="#ffc107")
        self.help_message.grid(row=0, column=0, sticky="e", pady=(0, 5))
        
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
        
        # Frame para tabla con scrollbar
        self.table_frame = ctk.CTkFrame(self.main_frame)
        self.table_frame.grid(row=2, column=0, sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(0, weight=1)
        
        # Frame para botones de admin (oculto por defecto)
        self.admin_buttons_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.admin_buttons_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.admin_buttons_frame.grid_remove()  # Ocultar inicialmente
        
        # Botones de administracion
        self.btn_agregar = ctk.CTkButton(self.admin_buttons_frame, text="➕ Agregar Cliente",
                                          command=self.mostrar_form_agregar,
                                          fg_color="#28a745", hover_color="#218838", width=150)
        self.btn_agregar.grid(row=0, column=0, padx=5)
        
        self.btn_modificar = ctk.CTkButton(self.admin_buttons_frame, text="✏️ Modificar",
                                            command=self.iniciar_modificacion,
                                            fg_color="#ffc107", hover_color="#e0a800", 
                                            text_color="black", width=150)
        self.btn_modificar.grid(row=0, column=1, padx=5)
        
        self.btn_eliminar = ctk.CTkButton(self.admin_buttons_frame, text="🗑️ Eliminar",
                                           command=self.eliminar_cliente,
                                           fg_color="#dc3545", hover_color="#c82333", width=150)
        self.btn_eliminar.grid(row=0, column=2, padx=5)
        
        self.btn_ver_todos = ctk.CTkButton(self.admin_buttons_frame, text="📋 Ver Todos",
                                            command=self.mostrar_todos_clientes,
                                            fg_color="#17a2b8", hover_color="#138496", width=150)
        self.btn_ver_todos.grid(row=0, column=3, padx=5)
        
        self.btn_guardar = ctk.CTkButton(self.admin_buttons_frame, text="💾 Actualizar y Guardar Cambios",
                                          command=self.guardar_cambios,
                                          fg_color="#007bff", hover_color="#0056b3", width=200)
        self.btn_guardar.grid(row=0, column=4, padx=(20, 5))
        
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
        columnas = ("id", "cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro")
        
        self.tabla = ttk.Treeview(self.table_frame, columns=columnas, show="headings", selectmode="browse")
        
        # Tag para hipervinculos (azul)
        self.tabla.tag_configure("link", foreground="#4da6ff")
        
        # Configurar columnas
        self.tabla.heading("id", text="ID")
        self.tabla.heading("cliente", text="Nombre Cliente")
        self.tabla.heading("ip_antena", text="IP Antena 🔗")
        self.tabla.heading("ip_router", text="IP Router 🔗")
        self.tabla.heading("ubicacion", text="Ubicacion")
        self.tabla.heading("zona", text="Zona")
        self.tabla.heading("plan", text="Plan")
        self.tabla.heading("fecha_registro", text="Fecha Reg.")
        
        # Ancho de columnas
        self.tabla.column("id", width=50, minwidth=40)
        self.tabla.column("cliente", width=150, minwidth=100)
        self.tabla.column("ip_antena", width=100, minwidth=80)
        self.tabla.column("ip_router", width=100, minwidth=80)
        self.tabla.column("ubicacion", width=120, minwidth=80)
        self.tabla.column("zona", width=100, minwidth=80)
        self.tabla.column("plan", width=80, minwidth=60)
        self.tabla.column("fecha_registro", width=100, minwidth=80)
        
        # Ocultar columna ID por defecto
        self.tabla.column("id", width=0, stretch=False)
        
        # Scrollbars
        scrollbar_y = ctk.CTkScrollbar(self.table_frame, command=self.tabla.yview)
        scrollbar_x = ctk.CTkScrollbar(self.table_frame, command=self.tabla.xview, orientation="horizontal")
        self.tabla.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Grid
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # Bind doble click para abrir IP o editar
        self.tabla.bind("<Double-1>", self.on_doble_click)
    
    def toggle_modo_admin(self):
        """Alternar entre modo normal y modo administrador"""
        self.modo_admin = not self.modo_admin
        
        if self.modo_admin:
            self.btn_admin.configure(text="🔙 Volver a Busqueda", fg_color="#6c757d")
            self.admin_buttons_frame.grid()
            self.titulo.configure(text="Administracion de Datos")
            # Mostrar columna ID
            self.tabla.column("id", width=50, stretch=True)
            # Ocultar seccion de zonas
            self.zonas_titulo.grid_remove()
            self.zonas_frame.grid_remove()
            self.separator2.grid_remove()
            # Mostrar todos los clientes
            self.mostrar_todos_clientes()
        else:
            self.btn_admin.configure(text="📝 Modificar datos BD", fg_color="#6c757d")
            self.admin_buttons_frame.grid_remove()
            self.titulo.configure(text="Sistema de Gestion DRECH")
            self.help_message.configure(text="")
            # Ocultar columna ID
            self.tabla.column("id", width=0, stretch=False)
            # Mostrar seccion de zonas
            self.zonas_titulo.grid()
            self.zonas_frame.grid()
            self.separator2.grid()
            # Limpiar tabla
            for item in self.tabla.get_children():
                self.tabla.delete(item)
    
    def mostrar_todos_clientes(self):
        """Mostrar todos los clientes en la tabla"""
        # Limpiar tabla
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        
        if not os.path.exists(DB_FILE):
            messagebox.showwarning("Aviso", "No hay base de datos. Carga un archivo Excel primero.")
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            
            # Verificar si existe columna rowid
            df = pd.read_sql_query("SELECT rowid as id, * FROM clientes", conn)
            conn.close()
            
            # Eliminar columnas unnamed
            df = df.loc[:, ~df.columns.str.contains('^unnamed', case=False)]
            
            for _, row in df.iterrows():
                valores = []
                for col in ["id", "cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro"]:
                    if col in df.columns:
                        val = row[col] if pd.notnull(row[col]) else ""
                        if col in ["ip_antena", "ip_router"] and val and str(val).strip():
                            val = f"🔗 {val}"
                        valores.append(val)
                    else:
                        valores.append("")
                self.tabla.insert("", "end", values=valores, tags=("link",))
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos: {e}")
    
    def mostrar_form_agregar(self):
        """Mostrar ventana para agregar nuevo cliente"""
        self.ventana_form = ctk.CTkToplevel(self)
        self.ventana_form.title("Agregar Nuevo Cliente")
        self.ventana_form.geometry("500x450")
        self.ventana_form.transient(self)
        self.ventana_form.grab_set()
        
        # Centrar ventana
        self.ventana_form.after(100, lambda: self.centrar_ventana(self.ventana_form))
        
        # Frame principal
        form_frame = ctk.CTkFrame(self.ventana_form)
        form_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Titulo
        ctk.CTkLabel(form_frame, text="Agregar Nuevo Cliente", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 20))
        
        # Campos
        campos = [
            ("Nombre Cliente:", "cliente"),
            ("IP Antena:", "ip_antena"),
            ("IP Router:", "ip_router"),
            ("Ubicacion:", "ubicacion"),
            ("Zona:", "zona"),
            ("Plan:", "plan"),
            ("Fecha Registro (DD-MM-YYYY):", "fecha_registro")
        ]
        
        self.entries_form = {}
        
        for label_text, field_name in campos:
            frame = ctk.CTkFrame(form_frame, fg_color="transparent")
            frame.pack(fill="x", pady=5)
            
            ctk.CTkLabel(frame, text=label_text, width=200, anchor="w").pack(side="left")
            entry = ctk.CTkEntry(frame, width=250, placeholder_text="DD-MM-YYYY" if field_name == "fecha_registro" else "")
            entry.pack(side="right", fill="x", expand=True)
            
            # Aplicar formato automatico de fecha
            if field_name == "fecha_registro":
                entry.bind("<KeyRelease>", formatear_fecha)
            
            self.entries_form[field_name] = entry
        
        # Botones
        btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancelar", command=self.ventana_form.destroy,
                      fg_color="#6c757d", width=100).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Guardar", command=self.guardar_nuevo_cliente,
                      fg_color="#28a745", width=100).pack(side="left", padx=10)
    
    def centrar_ventana(self, ventana):
        ventana.update_idletasks()
        width = ventana.winfo_width()
        height = ventana.winfo_height()
        x = (ventana.winfo_screenwidth() // 2) - (width // 2)
        y = (ventana.winfo_screenheight() // 2) - (height // 2)
        ventana.geometry(f'{width}x{height}+{x}+{y}')
    
    def guardar_nuevo_cliente(self):
        """Guardar nuevo cliente en la base de datos"""
        datos = {campo: entry.get().strip() for campo, entry in self.entries_form.items()}
        
        if not datos.get("cliente"):
            messagebox.showwarning("Aviso", "El nombre del cliente es obligatorio.")
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            
            # Obtener columnas existentes
            cursor.execute("PRAGMA table_info(clientes)")
            columnas_existentes = [col[1] for col in cursor.fetchall()]
            
            # Filtrar solo columnas que existen
            columnas = [c for c in datos.keys() if c in columnas_existentes]
            valores = [datos[c] for c in columnas]
            
            placeholders = ", ".join(["?" for _ in columnas])
            columnas_str = ", ".join(columnas)
            
            cursor.execute(f"INSERT INTO clientes ({columnas_str}) VALUES ({placeholders})", valores)
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Exito", "Cliente agregado correctamente.")
            self.ventana_form.destroy()
            self.actualizar_contador()
            self.mostrar_todos_clientes()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar cliente: {e}")
    
    def iniciar_modificacion(self):
        """Indicar al usuario que busque el cliente a modificar"""
        self.help_message.configure(text="⚠️ Busca aqui el usuario a modificar, luego haz doble clic en el para editarlo")
        self.search_entry.focus_set()
    
    def on_doble_click(self, event):
        """Manejar doble clic en la tabla"""
        item = self.tabla.selection()
        if not item:
            return
        
        region = self.tabla.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        col = self.tabla.identify_column(event.x)
        col_index = int(col.replace("#", "")) - 1
        
        # Si estamos en modo admin, abrir editor
        if self.modo_admin:
            self.editar_cliente(item[0])
        else:
            # Modo normal: abrir IP si es columna de IP
            if col_index in [2, 3]:  # ip_antena, ip_router (considerando ID oculto)
                valores = self.tabla.item(item[0], "values")
                ip = valores[col_index]
                if ip and ip.strip():
                    ip_limpia = ip.replace("🔗 ", "").strip()
                    if ip_limpia:
                        url = f"http://{ip_limpia}"
                        webbrowser.open(url)
    
    def editar_cliente(self, item_id):
        """Abrir ventana para editar cliente"""
        valores = self.tabla.item(item_id, "values")
        if not valores:
            return
        
        self.cliente_editando = valores[0]  # ID del cliente
        
        self.ventana_edit = ctk.CTkToplevel(self)
        self.ventana_edit.title("Editar Cliente")
        self.ventana_edit.geometry("500x450")
        self.ventana_edit.transient(self)
        self.ventana_edit.grab_set()
        
        self.ventana_edit.after(100, lambda: self.centrar_ventana(self.ventana_edit))
        
        form_frame = ctk.CTkFrame(self.ventana_edit)
        form_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(form_frame, text="Editar Cliente", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 20))
        
        campos = [
            ("Nombre Cliente:", "cliente", 1),
            ("IP Antena:", "ip_antena", 2),
            ("IP Router:", "ip_router", 3),
            ("Ubicacion:", "ubicacion", 4),
            ("Zona:", "zona", 5),
            ("Plan:", "plan", 6),
            ("Fecha Registro:", "fecha_registro", 7)
        ]
        
        self.entries_edit = {}
        
        for label_text, field_name, idx in campos:
            frame = ctk.CTkFrame(form_frame, fg_color="transparent")
            frame.pack(fill="x", pady=5)
            
            ctk.CTkLabel(frame, text=label_text, width=200, anchor="w").pack(side="left")
            entry = ctk.CTkEntry(frame, width=250, placeholder_text="DD-MM-YYYY" if field_name == "fecha_registro" else "")
            entry.pack(side="right", fill="x", expand=True)
            
            # Prellenar con valor actual (limpiar emoji si existe)
            valor_actual = valores[idx] if idx < len(valores) else ""
            valor_actual = str(valor_actual).replace("🔗 ", "")
            entry.insert(0, valor_actual)
            
            # Aplicar formato automatico de fecha
            if field_name == "fecha_registro":
                entry.bind("<KeyRelease>", formatear_fecha)
            
            self.entries_edit[field_name] = entry
        
        btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancelar", command=self.ventana_edit.destroy,
                      fg_color="#6c757d", width=100).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Guardar Cambios", command=self.guardar_edicion,
                      fg_color="#28a745", width=120).pack(side="left", padx=10)
    
    def guardar_edicion(self):
        """Guardar cambios del cliente editado"""
        datos = {campo: entry.get().strip() for campo, entry in self.entries_edit.items()}
        
        if not datos.get("cliente"):
            messagebox.showwarning("Aviso", "El nombre del cliente es obligatorio.")
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            
            # Obtener columnas existentes
            cursor.execute("PRAGMA table_info(clientes)")
            columnas_existentes = [col[1] for col in cursor.fetchall()]
            
            # Construir UPDATE
            updates = []
            valores = []
            for campo, valor in datos.items():
                if campo in columnas_existentes:
                    updates.append(f"{campo} = ?")
                    valores.append(valor)
            
            valores.append(self.cliente_editando)
            
            sql = f"UPDATE clientes SET {', '.join(updates)} WHERE rowid = ?"
            cursor.execute(sql, valores)
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Exito", "Cliente actualizado correctamente.")
            self.ventana_edit.destroy()
            self.mostrar_todos_clientes()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar cliente: {e}")
    
    def eliminar_cliente(self):
        """Eliminar cliente seleccionado"""
        item = self.tabla.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecciona un cliente de la tabla para eliminar.")
            return
        
        valores = self.tabla.item(item[0], "values")
        cliente_id = valores[0]
        cliente_nombre = valores[1].replace("🔗 ", "")
        
        confirmar = messagebox.askyesno("Confirmar", 
            f"¿Estas seguro de eliminar al cliente '{cliente_nombre}'?\nEsta accion no se puede deshacer.")
        
        if confirmar:
            try:
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM clientes WHERE rowid = ?", (cliente_id,))
                conn.commit()
                conn.close()
                
                messagebox.showinfo("Exito", "Cliente eliminado correctamente.")
                self.actualizar_contador()
                self.mostrar_todos_clientes()
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al eliminar cliente: {e}")
    
    def guardar_cambios(self):
        """Guardar todos los cambios y actualizar"""
        self.actualizar_contador()
        self.mostrar_todos_clientes()
        messagebox.showinfo("Exito", "Cambios guardados correctamente.\nLa base de datos ha sido actualizada.")
    
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
            df = df.loc[:, ~df.columns.str.contains('^unnamed', case=False)]
            
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
                # Actualizar zonas tambien
                self.actualizar_zonas()
            except:
                self.contador_label.configure(text="0")
        else:
            self.contador_label.configure(text="0")
            self.actualizar_zonas()
    
    def actualizar_zonas(self):
        """Actualizar la lista de zonas con sus cantidades de clientes"""
        # Limpiar botones existentes
        for widget in self.zonas_frame.winfo_children():
            widget.destroy()
        self.zonas_buttons = {}
        
        if not os.path.exists(DB_FILE):
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            # Obtener zonas y sus cantidades
            df = pd.read_sql_query("""
                SELECT zona, COUNT(*) as cantidad 
                FROM clientes 
                WHERE zona IS NOT NULL AND zona != ''
                GROUP BY zona 
                ORDER BY zona
            """, conn)
            conn.close()
            
            for _, row in df.iterrows():
                zona = row['zona']
                cantidad = row['cantidad']
                
                # Crear boton para cada zona
                btn = ctk.CTkButton(self.zonas_frame, 
                                    text=f"{zona} ({cantidad})",
                                    command=lambda z=zona: self.filtrar_por_zona(z),
                                    fg_color="transparent", hover_color="#3a3a3a",
                                    height=28, anchor="w",
                                    font=ctk.CTkFont(size=12))
                btn.pack(fill="x", pady=1, padx=5)
                self.zonas_buttons[zona] = btn
                
        except Exception as e:
            pass
    
    def filtrar_por_zona(self, zona):
        """Filtrar y mostrar clientes de una zona especifica"""
        self.zona_seleccionada = zona
        
        # Resaltar boton seleccionado
        for z, btn in self.zonas_buttons.items():
            if z == zona:
                btn.configure(fg_color="#1f6aa5")
            else:
                btn.configure(fg_color="transparent")
        
        # Limpiar tabla
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        
        if not os.path.exists(DB_FILE):
            return
        
        try:
            conn = sqlite3.connect(DB_FILE)
            df = pd.read_sql_query(f"""
                SELECT rowid as id, * FROM clientes 
                WHERE zona = ?
            """, conn, params=(zona,))
            conn.close()
            
            # Eliminar columnas unnamed
            df = df.loc[:, ~df.columns.str.contains('^unnamed', case=False)]
            
            for _, row in df.iterrows():
                valores = []
                for col in ["id", "cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro"]:
                    if col in df.columns:
                        val = row[col] if pd.notnull(row[col]) else ""
                        if col in ["ip_antena", "ip_router"] and val and str(val).strip():
                            val = f"🔗 {val}"
                        valores.append(val)
                    else:
                        valores.append("")
                self.tabla.insert("", "end", values=valores, tags=("link",))
            
            # Mostrar mensaje de filtro activo
            self.help_message.configure(text=f"📍 Mostrando zona: {zona} ({len(df)} clientes)")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al filtrar: {e}")
    
    def quitar_filtro_zona(self):
        """Quitar filtro de zona y limpiar tabla"""
        self.zona_seleccionada = None
        
        # Quitar resaltado de botones
        for btn in self.zonas_buttons.values():
            btn.configure(fg_color="transparent")
        
        # Limpiar tabla y mensaje
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        self.help_message.configure(text="")
    
    def buscar(self):
        query = self.search_entry.get().strip()
        
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
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(clientes)")
            columnas_disponibles = [col[1] for col in cursor.fetchall()]
            
            columnas_deseadas = ['cliente', 'ip_antena', 'ip_router', 'ubicacion', 'plan', 'fecha_registro', 'zona']
            columnas_select = [col for col in columnas_deseadas if col in columnas_disponibles]
            
            # Incluir rowid como id
            sql = f"""
            SELECT rowid as id, {', '.join(columnas_select)}
            FROM clientes 
            WHERE cliente LIKE '%{query}%' 
               OR ip_antena LIKE '%{query}%'
               OR ubicacion LIKE '%{query}%'
            """
            
            df = pd.read_sql_query(sql, conn)
            conn.close()
            
            if df.empty:
                messagebox.showinfo("Resultado", "No se encontraron coincidencias.")
                return
            
            for _, row in df.iterrows():
                valores = []
                for col in ["id", "cliente", "ip_antena", "ip_router", "ubicacion", "zona", "plan", "fecha_registro"]:
                    if col in df.columns:
                        val = row[col] if pd.notnull(row[col]) else ""
                        if col in ["ip_antena", "ip_router"] and val and str(val).strip():
                            val = f"🔗 {val}"
                        valores.append(val)
                    else:
                        valores.append("")
                self.tabla.insert("", "end", values=valores, tags=("link",))
                
        except Exception as e:
            messagebox.showerror("Error", f"Error en la busqueda: {e}")

if __name__ == "__main__":
    app = BuscadorDRECH()
    app.mainloop()
