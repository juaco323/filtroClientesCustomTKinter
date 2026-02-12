# Buscador Clientes DRECH - CustomTkinter

Sistema de gestión de clientes ISP con interfaz gráfica moderna usando CustomTkinter.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-5.2+-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## Características

- Cargar datos desde archivos Excel (.xlsx)
- Búsqueda por nombre de cliente, IP o ubicación
- **Sistema de filtros avanzados** por Ubicación y Zona
- Hipervínculos clickeables para IP Antena e IP Router
- **Validación de IPs** antes de abrir en navegador
- Base de datos SQLite local
- Interfaz moderna con tema oscuro
- Carga automática de ubicaciones y zonas al procesar Excel

## Capturas

La aplicación incluye:
- Panel lateral para carga de datos y contador de clientes
- Barra de búsqueda con filtros expandibles
- Tabla de resultados con columnas ordenables

## Requisitos

```bash
pip install customtkinter pandas openpyxl
```

## Uso

```bash
python app_ctk.py
```

### Flujo de trabajo

1. **Cargar Excel**: Selecciona el archivo Excel con los datos de clientes
2. **Procesar**: Presiona "Procesar y Actualizar BD" para importar los datos
3. **Buscar**: Ingresa nombre o IP en el campo de búsqueda
4. **Filtrar** (opcional): Presiona el botón "🔽 Filtrar" para mostrar filtros de Ubicación y Zona
5. **Ver resultados**: Los resultados aparecerán en la tabla
6. **Abrir IP**: Doble clic en una columna de IP para abrirla en el navegador

## Compilar a .exe

```bash
pip install pyinstaller
pyinstaller BuscadorDRECH_ctk.spec --clean --noconfirm
```

El ejecutable se generará en la carpeta `dist/`.

## Estructura del Excel

El archivo Excel debe tener las siguientes columnas:
- CLIENTE
- IP ANTENA
- IP ROUTER
- UBICACION
- PLAN
- Fecha Registro

Puede tener múltiples hojas, cada hoja se agregará como "zona".

## Tecnologías

- **Python 3.10+**
- **CustomTkinter** - Interfaz gráfica moderna
- **Pandas** - Procesamiento de datos Excel
- **SQLite3** - Base de datos local
- **ipaddress** - Validación de direcciones IP

## Autor

Desarrollado para DRECH ISP
