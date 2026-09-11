# Wardrobe Cutting Calculator v4.0

This package is the corrected standalone Python/Streamlit implementation of the `WD CUTTING` sheet in `Wardrobe cutting list.xlsx`.

It does **not** require Excel or LibreOffice.

## Run on Windows

1. Open PowerShell in this folder.
2. Install dependencies:

```powershell
py -m pip install -r requirements.txt
```

3. Run:

```powershell
py -m streamlit run app.py
```

### Sample verification
With the sample inputs from the workbook, rows 37–45 must be:

- BACK PANEL 1: 1990 x 991 x 1
- BACK PANEL 2: 1990 x 491 x 1
- BACK PANEL 3: 1990 x 0 x 0 (hidden because quantity is 0)
- DRAWER PACK: 483 x 200 x 1
- DR SIDE: 463 x 170 x 2
- DR BACK: 393 x 170 x 1
- DR IN FRONT: 393 x 50 x 1
- DR FRONT: 451 x 198 x 1
- BACK PANEL: 453 x 411 x 1
