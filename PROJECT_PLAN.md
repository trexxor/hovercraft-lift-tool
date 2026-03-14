# Hovercraft Lift Tool — V1 Planning Spec

## Purpose
Build a beginner-friendly, well-documented Excel VBA tool that estimates hovercraft lift power and supports scenario exploration with controlled solve-for behavior.

## Workflow Status
- Planning complete for V1 scope.
- No implementation code yet.

## Repository Guidance
- Keep this spec file in the root of the project repository (`hovercraft-lift-tool`).
- In future threads/sessions, share this file first so the implementation can resume from the same assumptions.

## V1 Scope
- Form-like Excel interface for naval architects/marine engineers.
- User-adjustable inputs with mode-based recalculation.
- Locked inputs (recalculation-immune unless user deliberately edits).
- Solve-for options for selected variables.
- Scenario logging: save current state to next row of a comparison table.
- Primary output: total lift power required.

## Global Toggles
- Unit system toggle: US Customary / SI.
- Pressure mode toggle: Absolute / Gauge.

## Key Naming and Unit Decisions
- Velocity uses lower-case symbols (`v_in`, `v_out`).
- Volumetric flow uses `Vdot` / `V̇`.
- US customary baseline uses slug-based mass units.

## Variables (V1)
- `p_a`: atmospheric pressure.
- `p_c`: plenum pressure.
- `Δp`: pressure differential.
- `ρ`: fluid density (locked input).
- `A_c`: effective plenum area.
- `W`: displaced vessel weight.
- `L`, `B`: ship length and beam.
- `h_gap`: air gap height.
- `A_j`: jet (air-gap) area.
- `ṁ`: mass flow rate.
- `v_in`, `v_out`: fan inflow and jet outflow velocities.
- `A_fan`, `r_fan`: fan disc area and fan radius (both editable + solve-for capable).
- `N_fans`: number of lift fans.
- `η_fan`: fan efficiency (locked input).
- `V̇`: volumetric flow rate.
- `F_lift`: lift force.
- `Weight-Lift Balance`: force difference check (`F_lift - W`).
- `P_lift`: required lift power.

## Core Equations (V1)
1. Pressure differential:
   - Absolute mode: `Δp = p_c - p_a`
   - Gauge mode: `Δp = p_c`
2. Lift force: `F_lift = Δp * A_c`
3. Weight-Lift Balance: `F_lift - W`
4. Jet continuity: `ṁ = ρ * A_j * v_out`
5. Fan continuity: `ṁ = ρ * A_fan * v_in`
6. Fan geometry: `A_fan = π * r_fan^2`
7. Volumetric flow: `V̇ = ṁ / ρ`
8. Lift power (initial model): `P_lift = (Δp * V̇) / η_fan`

## Behavior Rules
- Mode-based solving to avoid circular/ambiguous recalculation.
- Default sanity assumption: `v_in = v_out`.
- Flag warnings when equations imply a velocity mismatch.
- Locked inputs are never auto-overwritten during recalculation.
- Editing `A_fan` auto-updates `r_fan`; editing `r_fan` auto-updates `A_fan`.

## Scenario Logging (V1)
Each saved scenario appends:
- timestamp,
- scenario name/ID,
- unit/pressure mode,
- key inputs,
- key outputs,
- active solve modes,
- optional notes.

## V1 Exclusions
- Stability/trim/heel and CG-driven analysis.
- Advanced loss models and optimization.
- Final UI color behavior (deferred until after first working build).

## Code Style Requirements for Implementation
- Beginner-level VBA patterns.
- High comment density.
- Small modular procedures.
- Safe-to-edit structure to reduce regression risk.

## How to Reuse This Spec in a New Thread
1. Open the `hovercraft-lift-tool` repository.
2. Point Codex to `PROJECT_PLAN.md` first.
3. Ask Codex to implement from this file without changing locked assumptions unless explicitly requested.
