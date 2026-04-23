# Tessera
A Rhino 8 Python script that generates a parametric rectangular grid inside any closed planar curve. Useful for layout planning, paneling, fabrication grids, and similar workflows.

---

## Features

- **Two boundary input modes** — draw a new 3-point rectangle on the fly, or pick an existing closed curve from the scene
- **Full parametric control** — border, columns, rows, column gap, and row gap are all set interactively
- **Cell size preview** — shows calculated cell dimensions before committing, so you can adjust inputs if needed
- **Persistent settings** — last used values are saved to the document and pre-filled on next run
- **Auto layer setup** — creates a `Rectangle` parent layer (black) and `Rectangle::Cells` child layer (red) automatically
- **Cell labels** — optional A1/B2 style labels centered in each cell, auto-sized and placed on a separate `Rectangle::Labels` layer (blue)
- **Optional grouping** — all grid objects can be grouped under `GridGroup` for easy selection
- **Non-axis-aligned support** — works correctly on rotated rectangles and curves not aligned to the world XY plane, including Z-rotated shapes

---

## Requirements

- Rhino 8
- No external dependencies

---

## Running the Script

Download `Tessera.py` and save it somewhere permanent on your machine.

**Option 1 — RunPythonScript (no setup)**

Type in the Rhino command bar:
```
RunPythonScript
```
Browse to the file and open it. Simple, works every time.

**Option 2 — Alias (recommended)**

Run it by typing a short name in the command bar instead of browsing every time:

1. Go to `Tools > Options > Aliases`
2. Click `New`
3. Fill in:
   - **Name:** `Tessera` (or whatever you prefer)
   - **Command macro:** `-RunPythonScript "C:\path\to\Tessera.py"`
4. Click OK

From now on just type `Tessera` in the command bar and hit Enter.

**Option 3 — Toolbar Button**

1. Right-click on any toolbar > `New Button`
2. Paste this into the left mouse button command field:
   ```
   -RunPythonScript "C:\path\to\Tessera.py"
   ```
3. Add a label or icon if you want
4. Click OK

> Replace `C:\path\to\` with the actual folder where you saved the file.

---

## Usage

Run the script. It will walk you through the following steps:

**1. Choose boundary input method**
- `Draw3PointRectangle` — launches Rhino's native 3-point rectangle command
- `ExistingShape` — select any closed planar curve already in the scene

**2. Set parameters**
| Parameter | Description |
|---|---|
| Border | Inward offset from the boundary edge, equal on all sides |
| Columns | Number of columns |
| Rows | Number of rows |
| Gap X | Horizontal gap between columns |
| Gap Y | Vertical gap between rows |
| Cell labels | Optional Yes/No — adds A1/B2 style labels centered in each cell |

**3. Confirm cell size**
The script calculates and shows the resulting cell dimensions. You can cancel here and adjust inputs if the size isn't right.

**4. Choose grouping**
Optionally group all created objects under `GridGroup`.

**5. Summary**
```
Done.
-----------------------------
Boundary:    200 x 150 mm
Border:      5 mm
-----------------------------
Columns:     4
Rows:        3
Total cells: 12
-----------------------------
Cell size:   46.250 x 46.667 mm
Labels:      Yes
```

---

## Layers

| Layer | Color | Contents |
|---|---|---|
| `Rectangle` | Black | Outer boundary curve |
| `Rectangle::Cells` | Red | Cell rectangles |
| `Rectangle::Labels` | Blue | Cell labels (only if enabled) |

---

## Notes

- Re-running the script automatically deletes any previously created `GridGroup` before generating a new one
- Label format is column-letter + row-number: A1, A2, B1 ... Z1, AA1, AB1 ...
- Settings are saved per document via `SetDocumentData`, so they persist between runs in the same file

---

## Changelog

**v3.1**
- Added cell labels (A1/B2 format, auto-sized, separate layer)
- Fixed cell placement for rotated and Z-rotated curves
- Plane extraction now uses curve segments instead of CurvePlane

**v3.0**
- Added two boundary input modes (3-point rectangle / existing shape)
- Fixed cell placement bug for non-axis-aligned planes
- Cleaner summary screen with total cell count
- Removed redundant instruction prompt for 3-point rectangle

**v2.5**
- Initial public version