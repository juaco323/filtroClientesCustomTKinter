# Buscador Clientes DRECH - CustomTkinter

Sistema de gestion de clientes ISP con interfaz grafica usando CustomTkinter.

## Caracteristicas

- Cargar datos desde archivos Excel (.xlsx)
- Busqueda por nombre de cliente, IP o ubicacion
- Hipervinculos clickeables para IP Antena e IP Router
- Base de datos SQLite local
- Interfaz moderna con tema oscuro

## Requisitos

```
pip install customtkinter pandas openpyxl
```

## Uso

```
python app_ctk.py
```

## Compilar a .exe

```
pip install pyinstaller
pyinstaller BuscadorDRECH_ctk.spec --clean --noconfirm
```

El ejecutable se generara en la carpeta `dist/`.

## Estructura del Excel

El archivo Excel debe tener las siguientes columnas:
- CLIENTE
- IP ANTENA
- IP ROUTER
- UBICACION
- PLAN
- Fecha Registro

Puede tener multiples hojas, cada hoja se agregara como "zona".
