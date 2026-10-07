# FTIR KPI report

Notebook analysis of the FTIR extract in `data/`, plus an editable PowerPoint in `output/`.

## Python version

Use **Python 3.14.7, 64-bit Windows** (`win_amd64`).

The wheels in `packages/` match that interpreter. A 32-bit install, another Python minor version, or another operating system will not install them.

The Windows installer used for this project is:

`python-3.14.7-amd64.exe`

Download it once while online from:

https://www.python.org/ftp/python/3.14.7/python-3.14.7-amd64.exe

Copy that installer to the offline PC and install it there before the steps below. Tick **Add python.exe to PATH** during setup.

## Packages while you still have internet

From this folder:

```text
python download_packages.py
```

That runs `pip download` and stores every wheel, including dependencies, in `packages/`.

## Install with no internet

Copy the whole project folder, including `packages/`, `data/`, and `requirements.txt`. On the offline PC, with Python 3.14.7:

```text
python install_offline.py
```

That runs `pip install --no-index --find-links=packages -r requirements.txt`. Pip will not contact the network.

## Web app

Several people can open the site and download their own PowerPoint at the same time. The layout is in `ARCHITECTURE.md`.

```text
cd frontend
npm install
npm run build
cd ..
python run_web.py
```

Then open `http://127.0.0.1:8000`. `fastapi` and `uvicorn` are in `requirements.txt`. After you change those lines, run `python download_packages.py` again so the offline `packages/` folder includes them.

## Run the notebook

Open `ftir_kpi_report.ipynb` and run all cells.

The loader reads the first file it finds:

1. `data/ftir_dataset.csv`
2. `data/ftir_dataset.db`
3. `data/records.csv`

The deck is written to `output/FTIR_Automated_KPI_Report_2026.pptx`. If that file is open, the notebook writes `output/FTIR_Automated_KPI_Report_2026_updated.pptx` instead.

## Requirements

Pinned in `requirements.txt`:

- pandas 3.0.5
- numpy 2.5.2
- matplotlib 3.11.1
- python-pptx 1.0.2
- ipykernel 7.4.0 (so the notebook kernel can start offline)
