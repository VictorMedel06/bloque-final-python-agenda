# Agenda inteligente de contactos (Python)

Aplicación de escritorio para gestionar contactos con operaciones CRUD (crear, ver, actualizar, eliminar), búsqueda en tiempo real y herramientas de deduplicación.

## Librería externa

- `rapidfuzz`: se usa para:
  - Búsqueda inteligente (fuzzy) cuando escribes en el campo Buscar.
  - Detección de contactos duplicados por similitud de nombre (y coincidencia exacta por teléfono/email).

## Requisitos

- Python 3.9 o superior
- Tkinter (viene incluido con Python en Windows normalmente)

## Instalación

Instala la dependencia externa:

```bash
python -m pip install rapidfuzz
```

## Ejecución

Desde la carpeta del proyecto:

```bash
python main.py
```

Al ejecutarse, la aplicación crea una base de datos local `contacts.sqlite3` en la carpeta del proyecto. Este archivo no se sube al repositorio.

## Funcionalidades

- CRUD de contactos: añadir, actualizar, eliminar
- Búsqueda en tiempo real:
  - Exacta (sin RapidFuzz)
  - Fuzzy (con RapidFuzz instalado)
- Detección y fusión de duplicados (RapidFuzz)
- Importar / exportar contactos en CSV
- Ordenación por columnas (click en los encabezados)
- Modo oscuro (menú Ver → Modo oscuro o Ctrl+D)
- Menú contextual en la tabla (click derecho)

## Atajos de teclado

- Ctrl+F: enfocar búsqueda
- Ctrl+O: importar CSV
- Ctrl+S: exportar CSV
- Ctrl+D: alternar modo oscuro
- Enter: añadir/actualizar (si estás en los campos del formulario)
- Supr: eliminar (si la tabla está enfocada)
- Esc: limpiar formulario

## Estructura del proyecto

- `main.py`: arranque de la aplicación
- `contact_app/db.py`: acceso a SQLite (CRUD y consultas)
- `contact_app/ui.py`: interfaz Tkinter/ttk (ventana principal, menú, tabla, tooltips)
- `contact_app/validators.py`: validaciones (nombre, teléfono, email)
- `contact_app/duplicates.py`: deduplicación y fusión (RapidFuzz)
- `contact_app/models.py`: modelo `Contact`

