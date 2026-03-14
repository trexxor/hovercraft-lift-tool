Attribute VB_Name = "modHovercraftLiftTool"
Option Explicit

' ============================================================================
' Hovercraft Lift Tool (V1)
' ----------------------------------------------------------------------------
' Beginner-friendly VBA implementation based on PROJECT_PLAN.md.
' This module expects workbook-level named ranges for inputs and outputs.
'
' INPUT named ranges expected:
'   inp_pa, inp_pc, inp_rho, inp_Ac, inp_W, inp_L, inp_B, inp_h_gap,
'   inp_Aj, inp_mdot, inp_vin, inp_vout, inp_Afan, inp_rfan, inp_Nfans,
'   inp_eta_fan, inp_units_mode, inp_pressure_mode,
'   inp_solve_mode, inp_fan_geometry_driver, inp_scenario_name, inp_notes
'
' OUTPUT named ranges written:
'   out_dp, out_F_lift, out_balance, out_Vdot, out_P_lift, out_warning
' ============================================================================

Private Const PI_VAL As Double = 3.14159265358979
Private Const EPS As Double = 0.0000001

Public Sub RecalculateModel()
    ' Primary macro to recompute the model using current inputs.
    On Error GoTo CleanFail

    Dim m As HovercraftModel
    m = ReadModelFromSheet()

    ApplyFanGeometryDriver m
    SolveRequestedMode m
    CalculateOutputs m
    SyncFanGeometryPair m

    WriteModelOutputs m
    Exit Sub

CleanFail:
    MsgBox "RecalculateModel failed: " & Err.Description, vbExclamation, "Hovercraft Lift Tool"
End Sub

Public Sub SaveScenario()
    ' Appends the current state to a scenario log worksheet.
    On Error GoTo CleanFail

    Dim m As HovercraftModel
    m = ReadModelFromSheet()

    ApplyFanGeometryDriver m
    SolveRequestedMode m
    CalculateOutputs m

    AppendScenarioRow m
    WriteModelOutputs m
    Exit Sub

CleanFail:
    MsgBox "SaveScenario failed: " & Err.Description, vbExclamation, "Hovercraft Lift Tool"
End Sub

Public Sub InitializeDemoInputs()
    ' Optional helper for first-time setup/testing.
    SetNamedValue "inp_pa", 101325#
    SetNamedValue "inp_pc", 103500#
    SetNamedValue "inp_rho", 1.225
    SetNamedValue "inp_Ac", 120#
    SetNamedValue "inp_W", 250000#
    SetNamedValue "inp_L", 25#
    SetNamedValue "inp_B", 12#
    SetNamedValue "inp_h_gap", 0.08
    SetNamedValue "inp_Aj", 5.92
    SetNamedValue "inp_mdot", 220#
    SetNamedValue "inp_vin", 30#
    SetNamedValue "inp_vout", 30#
    SetNamedValue "inp_Afan", 4.52
    SetNamedValue "inp_rfan", 1.2
    SetNamedValue "inp_Nfans", 2
    SetNamedValue "inp_eta_fan", 0.75

    SetNamedValue "inp_units_mode", "SI"
    SetNamedValue "inp_pressure_mode", "ABSOLUTE"
    SetNamedValue "inp_solve_mode", "NONE"
    SetNamedValue "inp_fan_geometry_driver", "A_FAN"
    SetNamedValue "inp_scenario_name", "demo_case"
    SetNamedValue "inp_notes", ""

    RecalculateModel
End Sub

' ===========================
' Data model
' ===========================
Private Type HovercraftModel
    pa As Double
    pc As Double
    rho As Double
    Ac As Double
    W As Double
    L As Double
    B As Double
    h_gap As Double
    Aj As Double
    mdot As Double
    vin As Double
    vout As Double
    Afan As Double
    rfan As Double
    Nfans As Double
    eta_fan As Double

    units_mode As String
    pressure_mode As String
    solve_mode As String
    fan_geometry_driver As String
    scenario_name As String
    notes As String

    dp As Double
    F_lift As Double
    balance As Double
    Vdot As Double
    P_lift As Double
    warning_text As String
End Type

Private Function ReadModelFromSheet() As HovercraftModel
    Dim m As HovercraftModel

    m.pa = GetNamedDouble("inp_pa")
    m.pc = GetNamedDouble("inp_pc")
    m.rho = GetNamedDouble("inp_rho")
    m.Ac = GetNamedDouble("inp_Ac")
    m.W = GetNamedDouble("inp_W")
    m.L = GetNamedDouble("inp_L")
    m.B = GetNamedDouble("inp_B")
    m.h_gap = GetNamedDouble("inp_h_gap")
    m.Aj = GetNamedDouble("inp_Aj")
    m.mdot = GetNamedDouble("inp_mdot")
    m.vin = GetNamedDouble("inp_vin")
    m.vout = GetNamedDouble("inp_vout")
    m.Afan = GetNamedDouble("inp_Afan")
    m.rfan = GetNamedDouble("inp_rfan")
    m.Nfans = GetNamedDouble("inp_Nfans")
    m.eta_fan = GetNamedDouble("inp_eta_fan")

    m.units_mode = UCase$(GetNamedText("inp_units_mode"))
    m.pressure_mode = UCase$(GetNamedText("inp_pressure_mode"))
    m.solve_mode = UCase$(GetNamedText("inp_solve_mode"))
    m.fan_geometry_driver = UCase$(GetNamedText("inp_fan_geometry_driver"))
    m.scenario_name = GetNamedText("inp_scenario_name")
    m.notes = GetNamedText("inp_notes")

    ReadModelFromSheet = m
End Function

' ===========================
' Solve logic
' ===========================
Private Sub SolveRequestedMode(ByRef m As HovercraftModel)
    Select Case m.solve_mode
        Case "NONE", ""
            ' No solve-for action.

        Case "SOLVE_PC_FOR_WEIGHT"
            ' Target condition: F_lift = W, where F_lift = dp * Ac
            If m.Ac <= EPS Then Err.Raise vbObjectError + 8001, , "Cannot solve pc: Ac must be > 0."

            If m.pressure_mode = "GAUGE" Then
                m.pc = m.W / m.Ac
            Else
                m.pc = (m.W / m.Ac) + m.pa
            End If

        Case "SOLVE_AC_FOR_WEIGHT"
            ' Target condition: Ac = W / dp
            m.dp = ComputeDeltaP(m.pc, m.pa, m.pressure_mode)
            If m.dp <= EPS Then Err.Raise vbObjectError + 8002, , "Cannot solve Ac: Δp must be > 0."
            m.Ac = m.W / m.dp

        Case "SOLVE_VOUT_FOR_MDOT"
            ' Target condition: vout = mdot / (rho * Aj)
            If m.rho <= EPS Or m.Aj <= EPS Then Err.Raise vbObjectError + 8003, , "Cannot solve v_out: rho and Aj must be > 0."
            m.vout = m.mdot / (m.rho * m.Aj)

        Case "SOLVE_AFAN_FOR_MDOT"
            ' Target condition: Afan = mdot / (rho * vin)
            If m.rho <= EPS Or m.vin <= EPS Then Err.Raise vbObjectError + 8004, , "Cannot solve A_fan: rho and v_in must be > 0."
            m.Afan = m.mdot / (m.rho * m.vin)

        Case "SOLVE_RFAN_FOR_AFAN"
            ' Target condition: rfan = sqrt(Afan / pi)
            If m.Afan <= EPS Then Err.Raise vbObjectError + 8005, , "Cannot solve r_fan: A_fan must be > 0."
            m.rfan = Sqr(m.Afan / PI_VAL)

        Case "SOLVE_AFAN_FOR_RFAN"
            ' Target condition: Afan = pi * rfan^2
            If m.rfan <= EPS Then Err.Raise vbObjectError + 8006, , "Cannot solve A_fan: r_fan must be > 0."
            m.Afan = PI_VAL * m.rfan * m.rfan

        Case Else
            m.warning_text = AppendWarning(m.warning_text, "Unknown solve mode: " & m.solve_mode)
    End Select
End Sub

Private Sub ApplyFanGeometryDriver(ByRef m As HovercraftModel)
    ' Behavior rule: editing A_fan updates r_fan; editing r_fan updates A_fan.
    Select Case m.fan_geometry_driver
        Case "A_FAN", ""
            If m.Afan > EPS Then m.rfan = Sqr(m.Afan / PI_VAL)

        Case "R_FAN"
            If m.rfan > EPS Then m.Afan = PI_VAL * m.rfan * m.rfan

        Case Else
            m.warning_text = AppendWarning(m.warning_text, "Unknown fan geometry driver: " & m.fan_geometry_driver)
    End Select
End Sub

Private Sub SyncFanGeometryPair(ByRef m As HovercraftModel)
    ' Force final consistency between A_fan and r_fan after calculations.
    If m.rfan > EPS Then
        m.Afan = PI_VAL * m.rfan * m.rfan
    ElseIf m.Afan > EPS Then
        m.rfan = Sqr(m.Afan / PI_VAL)
    End If

    ' Write back synchronized values to input cells.
    SetNamedValue "inp_Afan", m.Afan
    SetNamedValue "inp_rfan", m.rfan
End Sub

' ===========================
' Core calculations
' ===========================
Private Sub CalculateOutputs(ByRef m As HovercraftModel)
    m.dp = ComputeDeltaP(m.pc, m.pa, m.pressure_mode)

    ' F_lift = Δp * A_c
    m.F_lift = m.dp * m.Ac

    ' Force difference check.
    m.balance = m.F_lift - m.W

    ' Continuity calculations.
    Dim mdot_from_jet As Double
    Dim mdot_from_fan As Double

    mdot_from_jet = m.rho * m.Aj * m.vout
    mdot_from_fan = m.rho * m.Afan * m.vin

    ' Sanity assumption from spec: v_in = v_out.
    If Abs(m.vin - m.vout) > 0.001 Then
        m.warning_text = AppendWarning(m.warning_text, "v_in and v_out differ from the default assumption.")
    End If

    ' If user-provided mdot is near-zero, backfill it from jet continuity.
    If m.mdot <= EPS Then
        m.mdot = mdot_from_jet
    End If

    If Abs(m.mdot - mdot_from_jet) > 0.01 * MaxNonZero(m.mdot, mdot_from_jet) Then
        m.warning_text = AppendWarning(m.warning_text, "ṁ mismatch vs jet continuity.")
    End If

    If Abs(m.mdot - mdot_from_fan) > 0.01 * MaxNonZero(m.mdot, mdot_from_fan) Then
        m.warning_text = AppendWarning(m.warning_text, "ṁ mismatch vs fan continuity.")
    End If

    ' Vdot = mdot / rho
    If m.rho <= EPS Then Err.Raise vbObjectError + 8007, , "rho must be > 0 for volumetric flow."
    m.Vdot = m.mdot / m.rho

    ' P_lift = (Δp * Vdot) / eta_fan
    If m.eta_fan <= EPS Then Err.Raise vbObjectError + 8008, , "eta_fan must be > 0 for power calculation."
    m.P_lift = (m.dp * m.Vdot) / m.eta_fan
End Sub

Private Function ComputeDeltaP(ByVal pc As Double, ByVal pa As Double, ByVal pressureMode As String) As Double
    Select Case UCase$(pressureMode)
        Case "GAUGE"
            ComputeDeltaP = pc
        Case Else
            ComputeDeltaP = pc - pa
    End Select
End Function

Private Sub WriteModelOutputs(ByVal m As HovercraftModel)
    SetNamedValue "out_dp", m.dp
    SetNamedValue "out_F_lift", m.F_lift
    SetNamedValue "out_balance", m.balance
    SetNamedValue "out_Vdot", m.Vdot
    SetNamedValue "out_P_lift", m.P_lift
    SetNamedValue "out_warning", m.warning_text
End Sub

' ===========================
' Scenario logging
' ===========================
Private Sub AppendScenarioRow(ByVal m As HovercraftModel)
    Dim ws As Worksheet
    Set ws = EnsureScenarioSheet()

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = m.scenario_name
    ws.Cells(nextRow, 3).Value = m.units_mode
    ws.Cells(nextRow, 4).Value = m.pressure_mode
    ws.Cells(nextRow, 5).Value = m.solve_mode

    ws.Cells(nextRow, 6).Value = m.pa
    ws.Cells(nextRow, 7).Value = m.pc
    ws.Cells(nextRow, 8).Value = m.rho
    ws.Cells(nextRow, 9).Value = m.Ac
    ws.Cells(nextRow, 10).Value = m.W
    ws.Cells(nextRow, 11).Value = m.Aj
    ws.Cells(nextRow, 12).Value = m.mdot
    ws.Cells(nextRow, 13).Value = m.vin
    ws.Cells(nextRow, 14).Value = m.vout
    ws.Cells(nextRow, 15).Value = m.Afan
    ws.Cells(nextRow, 16).Value = m.rfan
    ws.Cells(nextRow, 17).Value = m.Nfans
    ws.Cells(nextRow, 18).Value = m.eta_fan

    ws.Cells(nextRow, 19).Value = m.dp
    ws.Cells(nextRow, 20).Value = m.F_lift
    ws.Cells(nextRow, 21).Value = m.balance
    ws.Cells(nextRow, 22).Value = m.Vdot
    ws.Cells(nextRow, 23).Value = m.P_lift
    ws.Cells(nextRow, 24).Value = m.warning_text
    ws.Cells(nextRow, 25).Value = m.notes
End Sub

Private Function EnsureScenarioSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ScenarioLog")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "ScenarioLog"
    End If

    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Range("A1:Y1").Value = Array( _
            "timestamp", "scenario_name", "units_mode", "pressure_mode", "solve_mode", _
            "p_a", "p_c", "rho", "A_c", "W", "A_j", "mdot", "v_in", "v_out", _
            "A_fan", "r_fan", "N_fans", "eta_fan", _
            "delta_p", "F_lift", "weight_lift_balance", "Vdot", "P_lift", "warnings", "notes")
        ws.Rows(1).Font.Bold = True
    End If

    Set EnsureScenarioSheet = ws
End Function

' ===========================
' Workbook named-range utilities
' ===========================
Private Function GetNamedDouble(ByVal rangeName As String) As Double
    GetNamedDouble = CDbl(ThisWorkbook.Names(rangeName).RefersToRange.Value)
End Function

Private Function GetNamedText(ByVal rangeName As String) As String
    GetNamedText = CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value)
End Function

Private Sub SetNamedValue(ByVal rangeName As String, ByVal valueIn As Variant)
    ThisWorkbook.Names(rangeName).RefersToRange.Value = valueIn
End Sub

Private Function AppendWarning(ByVal existingText As String, ByVal newText As String) As String
    If Trim$(existingText) = "" Then
        AppendWarning = newText
    Else
        AppendWarning = existingText & " | " & newText
    End If
End Function

Private Function MaxNonZero(ByVal a As Double, ByVal b As Double) As Double
    MaxNonZero = Abs(a)
    If Abs(b) > MaxNonZero Then MaxNonZero = Abs(b)
    If MaxNonZero < EPS Then MaxNonZero = 1#
End Function
