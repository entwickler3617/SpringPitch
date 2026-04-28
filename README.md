# Spring Reverse Engineering Program (Steps 1–4)

This repository implements the analysis pipeline for measuring spring geometry from IGS point data, delivering all outputs for steps 1–4. Step 5 (three‑point arc precision metrics) is intentionally excluded per request.

## What’s implemented ✅

- IGS parsing and PCA-based axis detection (principal axis and center)
- Coordinate transforms: xyz → local (x′,y′,z′) → cylindrical (R, θ, Z)
- Standardization to 1,000 points along the spring path (choose method via env)
  - Linear (arc-length polyline)
  - Uniform B‑Spline (default; SciPy splprep/splev)
  - NURBS (geomdl) with input decimation and graceful fallback
- Origin and θ start determination
  - Start/end slice chosen by "min median radius" at ends (robust against noise)
  - θ shifted globally so that it starts at 0 from the origin point
  - Final origin = axis-projected midpoint of start and opposite-phase (θ≈π) points
- Angle handling and display
  - θ_unwrapped (analysis), θ wrapped to −π..π (rad), θ wrapped to −180..180 (deg)
  - Excel displays θ in degrees, wrapped (−180..180)
- Basic metrics (step 4): turn, height, (local) pitch, smoothed R/Z, diameters
- Excel writer (sheet `zero-1`) that safely updates only target columns
- Visualization helpers and Excel plotter

Step 5 (three-point arc precision) is not produced by design.

## Conversion Steps (Step1 → Step3)

This pipeline writes all original IGS points to Excel and derives normalized and cylindrical coordinates from the same points. In the provided sample, Step1 contains 3141 rows.

Sheet mapping:

- Step1 → Excel sheet `raw` (columns: No, x, y, z; rows = all original points)
- Step2/Step3 → Excel sheet `zero-1` (1,000 standardized rows; columns E:J)

- Step1 (sheet `raw`): (x, y, z)
  - All original IGS coordinates in millimeters (type-116 entities).
  - IGES 124 transforms applied; units normalized to mm.
  - Written to `TK1.xlsx` → sheet `raw` columns B:D starting at row 2 (with `No` in A).

- Step2 (sheet `zero-1`): (x_norm, y_norm, z_norm)
  - The path is ordered and resampled to exactly 1,000 points with equal arc-length spacing (equidistant along curve).
  - PCA finds spring axis; local frame aligns the axis to local X.
  - Direction flipped if needed so X increases; `x_norm` starts at 0.
  - `y_norm, z_norm` centered by subtracting their mean (per-run dataset mean).
  - Written to `zero-1` columns E:G (1,000 rows).

- Step3 (sheet `zero-1`): (R, θ, Z)
  - Cylinder defined along local X; radial plane is local YZ.
  - `Z = x_norm`.
  - `R = sqrt(y_norm^2 + z_norm^2)`.
  - `θ = atan2(z_norm, y_norm)`; unwrapped and shifted so the first point starts at 0°, then wrapped for display where needed.
  - Written to `zero-1` columns H:J (R, θ in degrees, Z in mm).

Notes:

- In Excel, Step1 columns B:D are the nearest raw coordinates to each standardized row, keeping the table aligned; the full original raw list (e.g., 3141 points) is exported to `output/points_xyz.csv`.
- Auxiliary L:X columns (turns, diameters, pitch, etc.) are computed from the Step3/θ sequence; headers are hardcoded to match the std spec; atomic saves prevent file corruption.
 - Geometry preservation: Step2's 1,000 rows are equal-arc-length samples along the original path (interpolated), so the underlying geometry is not altered; the original raw dataset remains unchanged and available.
 - Row counts: Excel Step2/Step3 contain exactly 1,000 rows; Excel Step1 shows the mapped nearest raw points for those rows. The complete raw set is not truncated—see the CSV dump.
- Raw export and preview:

```powershell
# Export full raw list and a quick 3D preview image
F:/SpringPitch/.venv/Scripts/python.exe f:/SpringPitch/tools/export_xyz_and_plot.py
```

- If you prefer a separate Excel sheet with all raw points (e.g., 3141 rows), I can add an optional `raw` sheet writer—just say the word.

## Project structure

- `detailed_center_analysis.py` — main pipeline (IGS → analysis → CSV/Excel/plots + new result report chart)
- `plot_tk1_excel.py` — reads `TK1.xlsx` (sheet `zero-1`) and renders figures
- `final_report_generator.py` — creates comprehensive final report chart
- `output/` — all generated CSVs, summary, Excel (also a copy at repo root)
- `requirements_analysis.md` — spec clarifications and decisions
- `TK1_FRT_zero-1_251014.igs` — sample input IGS

## How to run

PowerShell (Windows):

```powershell
# 1) Run the main analysis (writes CSVs/summary/Excel and overview plot)
F:/SpringPitch/.venv/Scripts/python.exe f:/SpringPitch/detailed_center_analysis.py

# 2) (Optional) Plot from the generated Excel (produces PNGs under output/)
F:/SpringPitch/.venv/Scripts/python.exe f:/SpringPitch/plot_tk1_excel.py
```

### Standalone EXE (Windows)

You can run the packaged executable without Python installed.

Placement options:

- Put `SpringCalculator.exe`, your `.igs` file, and `TK1.xlsx` in the same folder and double‑click the EXE.
- Or pass the IGES path explicitly from PowerShell:

```powershell
# Run from anywhere with explicit IGES path
F:/SpringPitch/dist/SpringCalculator.exe F:/SpringPitch/TK1_FRT_zero-1_251014.igs
```

Behavior:

- The app searches for the IGES in this order: CLI arg → EXE folder → parent folder → current working directory.
- `TK1.xlsx` is looked up beside the IGES first; if missing, it falls back to the EXE folder, its parent, then CWD.
- On each run, `zero-1` sheet rows ≥2 in columns A:X are cleared before new data is written; formatting is preserved.
- Step1 full raw points go to sheet `raw`; Step2/3 standardized 1,000 rows go to `zero-1`.
- Images like `spring_detailed_analysis.png` and `new_result_report.png` are saved next to the IGES file.

### Configuration

Use an environment variable to select the standardization method:

```powershell
$env:SPRING_NORM_METHOD = "uniform_bspline"  # linear | uniform_bspline | nurbs
F:/SpringPitch/.venv/Scripts/python.exe f:/SpringPitch/detailed_center_analysis.py
```

Defaults:

- Method: `uniform_bspline`
- B‑spline degree: 3 (auto-limited by points available)
- B‑spline smoothing: 0.0 (interpolation-like; >0 applies smoothing)

## Units

All dimensional values use millimeters (mm) and degrees for angular display. Mapping:

| Column / Metric            | Meaning                              | Unit |
|----------------------------|--------------------------------------|------|
| x, y, z (raw)              | Original IGES coordinates            | mm   |
| x_norm, y_norm, z_norm     | Raw-preserved (or lightly corrected) | mm   |
| R                          | Radius (sqrt(x_norm^2+y_norm^2))     | mm   |
| θ (sheet)                  | Wrapped angle (starts at 0)          | deg  |
| Z                          | Axial coordinate (origin-shifted)    | mm   |
| Turn                       | Cumulative revolutions (θ/360)       | rev  |
| Diameter (Perpendicular)   | 2*R                                  | mm   |
| Diameter (Vectorial)       | Distance to indexed raw point        | mm   |
| Height                     | Z - Z(start)                         | mm   |
| Pitch                      | Axial advance since previous turn    | mm   |
| min_Pitch                  | Minimum pitch over column W          | mm   |
| L_delta_theta              | Change in wrapped theta              | deg  |
| M_cum_theta                | Cumulative theta changes             | deg  |
| N_turn                     | Turn (θ_unwrapped/360)               | rev  |
| O_radius_copy              | Copy of radius                       | mm   |
| P_half_turn_idx            | Index for half-turn back (MATCH)     | -    |
| Q_perp_diam                | Perpendicular diameter sum           | mm   |
| R_fallback_diam            | Fallback vectorial diameter          | mm   |
| S_vec_diam                 | Vectorial diameter                   | mm   |
| T_abs_z                    | Absolute Z coordinate                | mm   |
| U_rel_height               | Relative height from start           | mm   |
| V_full_turn_idx            | Index for full-turn back             | -    |
| W_pitch                    | Pitch calculation                    | mm   |
| X_min_pitch                | Minimum pitch value                  | mm   |

## Outputs

Files written under `output/` (and an Excel copy at repository root):

- `rtheta_z_table.csv`
  - x_local, y_local, z_local, r, theta, theta_unwrapped, turn
- `dimensions_basic.csv`
  - z_local, r_smooth, diameter_perpendicular, diameter_vectorial, turn, pitch_local, pitch_smooth
- `summary_basic.txt`
  - height_total, pitch_mean, pitch_min_estimate, turn_total_estimate
- `standardized_table.csv`
  - No, x, y, z, x_norm, y_norm, z_norm, R,
  - theta_unwrapped_rad, theta_wrap_rad, theta_wrap_deg, Z
- `TK1.xlsx` (also a copy at `f:/SpringPitch/TK1.xlsx` if not locked)
  - Sheet: `zero-1`
  - Headers in row 1; data from row 2
  - Core columns: No, x, y, z, x_norm, y_norm, z_norm, R, θ, Z
  - Auxiliary & derived (L..X region) implement Excel formula equivalents: delta theta, cumulative theta, turn, radius copy, half-turn index (MATCH), perpendicular diameter, vectorial diameter, absolute/relative Z, full-turn index, pitch, and min pitch.

Overview visualization:

- `spring_detailed_analysis.png` — PCA axis, layer summaries, multi-view plots
- `final_report.png` — Comprehensive summary chart with 3D visualization, key metrics, and analysis plots
- `new_result_report.png` — Detailed result report chart based on Z:AT column data with multi-panel analysis

If you run `plot_tk1_excel.py`, additional figures are saved to `output/`:

- `xstd_ystd_zstd_3d.png` — 3D trajectory of standardized points
- `xstd_ystd_zstd_vs_index.png` — x_std, y_std, z_std vs index
- `R_theta_Z_vs_index.png` — R, θ (deg, wrapped), Z vs index

## Method details

### Axis and frame

- PCA on xyz points → principal axis (Ez) and mean center
- Build orthonormal frame (Ex,Ey,Ez); local coords are world points shifted by origin and rotated by this frame

### Origin and start index

1) Partition ends by axis projection (e.g., bottom/top 5%).

2) Pick the end with smaller median radius as the start slice (robust to outliers).

3) Provisional origin = axis projection of the start-slice center.

4) Determine `start_idx` as the standardized point closest to the start-slice center; compute opposite-phase index (θ≈π) using unwrapped θ referenced to `start_idx`.

5) Final origin = axis projection of the midpoint between start and opposite-phase points.

“Option A”: arrays are NOT rolled. We keep the original sequence to avoid Z seam artifacts; θ is zero-biased at `start_idx` for interpretability.

### Standardization (1,000 points)

- Linear: arc-length resampling along the input polyline
- Uniform B‑Spline: SciPy `splprep/splev` with uniform parameter sampling
- NURBS: geomdl `fitting.approximate_curve`; input decimated if too dense
- In all cases, nearest original indices are tracked for raw↔std mapping

### Angles

- Store unwrapped θ for analysis; wrap to −π..π for rad and −180..180 for deg
- Excel writes θ as degrees (wrapped) per display spec; CSV includes all variants

## Dependencies

- Python 3.13 (venv used in this repo)
- numpy, matplotlib, openpyxl
- SciPy (for B‑Spline); geomdl (optional, for NURBS)

If NURBS or SciPy are unavailable, the code falls back to a supported method and logs a warning.

## Known limitations / next steps

- Step 5 precision metrics (three‑point arc) are intentionally excluded
- Optional knobs to expose (can be added on request):
  - End-slice percent, B‑spline smoothing/degree, NURBS decimation size
  - Optional continuous Z metric column

## Changelog (highlights)

- **v0.9 (November 2025)**: Updated cylindrical coordinate conversion to ensure theta starts at 0 from the origin point. Implemented Excel formula equivalents for additional columns (O: radius copy, P: half-turn index via MATCH, Q: perpendicular diameter sums, R: vectorial diameter fallback, S: vectorial diameter, T: absolute Z, U: relative height, V: full-turn index, W: pitch, X: min pitch). Fixed turn calculations using unwrapped theta. Added file copying to ensure TK1.xlsx is placed in output folder. Enhanced outlier detection and curvature-based analysis.
- Added switchable standardization (linear | uniform_bspline | nurbs)
- Implemented robust origin and θ start selection; removed array rolling (Option A)
- θ handling improvements: zero-bias, unwrap, and wrapped display (deg)
- Excel writer hardened to safely update only target columns
- Added plotting script to visualize standardized outputs
