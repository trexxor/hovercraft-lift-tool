# Hovercraft Lift Tool (V1 Implementation Scaffold)

This repository now includes a beginner-friendly VBA module implementing the V1 equations and behaviors from `PROJECT_PLAN.md`.

## Files

- `PROJECT_PLAN.md` — V1 planning specification (source of truth).
- `vba/modHovercraftLiftTool.bas` — importable VBA module with calculation engine and scenario logging.

## How to use in Excel

1. Open your workbook (`.xlsm`) and press `ALT+F11` to open the VBA editor.
2. Import `vba/modHovercraftLiftTool.bas`.
3. Create workbook-level named ranges for the expected input and output names (see comments at top of module).
4. Add buttons (optional) that call:
   - `InitializeDemoInputs`
   - `RecalculateModel`
   - `SaveScenario`

## Design notes

- Beginner-friendly and heavily commented procedures.
- Locked-input principle is respected by not auto-writing `inp_rho` or `inp_eta_fan`.
- Fan geometry two-way behavior is controlled by `inp_fan_geometry_driver`:
  - `A_FAN`: recompute `r_fan` from `A_fan`.
  - `R_FAN`: recompute `A_fan` from `r_fan`.
- Pressure mode supported via `inp_pressure_mode`:
  - `ABSOLUTE`: `Δp = p_c - p_a`
  - `GAUGE`: `Δp = p_c`

## Solve mode values supported

- `NONE`
- `SOLVE_PC_FOR_WEIGHT`
- `SOLVE_AC_FOR_WEIGHT`
- `SOLVE_VOUT_FOR_MDOT`
- `SOLVE_AFAN_FOR_MDOT`
- `SOLVE_RFAN_FOR_AFAN`
- `SOLVE_AFAN_FOR_RFAN`
