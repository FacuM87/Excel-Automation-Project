# MisLiquidaciones

Aplicación de escritorio para Windows que arma la liquidación mensual de prestaciones a partir del Excel de órdenes recibido (`dgOrdenes.XLSX`).

- **Liquidar**: lee el archivo, aplica el monto de cada prestación y genera el Excel de liquidación con el total.
- **Historial**: guarda cada liquidación en una base local y permite importar archivos de meses anteriores.
- **Reportes**: totales mensuales y anuales por prestación, con gráficos y exportación a Excel.
- **Configuración**: montos por prestación, carpeta de guardado y tema claro u oscuro.

## Tecnologías

- Python 3.12
- [pandas](https://pandas.pydata.org/) y [openpyxl](https://openpyxl.readthedocs.io/) para leer y escribir Excel
- [ttkbootstrap](https://ttkbootstrap.readthedocs.io/) (Tkinter) para la interfaz
- SQLite para el historial
- [PyInstaller](https://pyinstaller.org/) para generar el `.exe`
- pytest para los tests

## Descargar el código

```bash
git clone https://github.com/FacuM87/Excel-Automation-Project.git
cd Excel-Automation-Project
```

## Ejecutar localmente

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Para correr los tests:

```bash
pip install -r requirements-dev.txt
pytest
```

## Generar el .exe

```bash
pyinstaller MisLiquidaciones.spec --clean
```

El ejecutable queda en `dist\MisLiquidaciones.exe`. Guarda sus datos (montos, ajustes e historial) en `%APPDATA%\MisLiquidaciones`, así que se puede mover o reemplazar sin perderlos.
