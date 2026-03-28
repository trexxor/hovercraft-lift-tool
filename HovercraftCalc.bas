Attribute VB_Name = "HovercraftCalc"
' ============================================================
' MODULE: HovercraftCalc
'
' PURPOSE:
'   This is the main calculation engine for the Hovercraft
'   Performance Calculator. It contains all the physics
'   formulas and the logic that decides WHICH formula to use
'   depending on which variable the user has selected to
'   solve for (the "solve for" radio button on the form).
'
' HOW THIS MODULE FITS IN:
'   HovercraftCalc.bas  <- YOU ARE HERE (calculation logic)
'   frmHovercraft.frm   <- The UserForm (visual interface)
'   ThisWorkbook.bas    <- Opens the form when the file loads
'
' UNIT CONVENTIONS USED IN ALL CALCULATIONS:
'   Pressure  : lb/ft²  (converted to/from psia/psig for display)
'   Length    : ft
'   Area      : ft²
'   Weight    : lb      (converted to/from Long Tons for display)
'   Density   : slug/ft³
'   Flow      : ft³/s (volumetric),  slug/s (mass)
'   Power     : hp     (1 hp = 550 ft·lb/s)
'   Velocity  : ft/s
'
' CONVERSION CONSTANTS:
'   1 Long Ton  = 2240 lb
'   1 psi       = 144 lb/ft²  (1 ft² = 144 in²)
'   1 hp        = 550 ft·lb/s
' ============================================================

Option Explicit

' --- Module-level conversion constants ---
' "Public Const" makes these available to all other modules too.
' "Const" means the value never changes while the program runs.
Public Const LB_PER_LT      As Double = 2240#               ' pounds per Long Ton
Public Const IN2_PER_FT2    As Double = 144#                ' square inches per square foot
Public Const FTLBS_PER_HP   As Double = 550#                ' ft·lb/s per horsepower
Public Const PI             As Double = 3.14159265358979

' --- Module-level unit system state and dictionaries ---
' isUS is the single source of truth for the current unit system.
' True = US Customary is currently displayed.
' False = SI is currently displayed.
' Initialized to True in UserForm_Initialize.
Public isUS As Boolean

' Four dictionaries that store current values and unit strings.
' Populated by PopulateDictionaries after every calculation.
' Read by DisplayFromDictionary when updating the form display.
Public dictUSValues As Scripting.Dictionary
Public dictSIValues As Scripting.Dictionary
Public dictUSUnits  As Scripting.Dictionary
Public dictSIUnits  As Scripting.Dictionary

' --- Module-level calculation variables ---
' Declared here so PopulateDictionaries can access them without parameters.
' All values are in US Customary base units at all times.
Private p_a_psi     As Double   ' Ambient pressure, psia
Private rho         As Double   ' Air density, slug/ft³
Private W_ls_LT     As Double   ' Lightship weight, Long Tons
Private W_dw_LT     As Double   ' Deadweight, Long Tons
Private W_tot_LT    As Double   ' Total weight, Long Tons
Private L_c         As Double   ' Cushion length, ft
Private B_c         As Double   ' Cushion beam, ft
Private A_c         As Double   ' Cushion plan area, ft²
Private D_sk        As Double   ' Skirt depth, ft
Private D_ag        As Double   ' Air gap depth, ft
Private D_c         As Double   ' Cushion depth, ft
Private eta         As Double   ' Fan efficiency, dimensionless
Private P_f_hp      As Double   ' Fan power, hp
Private C_v         As Double   ' Coefficient of velocity, dimensionless
Private C_c         As Double   ' Coefficient of contraction, dimensionless
Private C_d         As Double   ' Discharge coefficient, dimensionless
Private theta       As Double   ' Skirt fingers angle at ground, degrees
Private p_a         As Double   ' Ambient pressure, lb/ft²
Private W_ls        As Double   ' Lightship weight, lb
Private W_dw        As Double   ' Deadweight, lb
Private W_tot       As Double   ' Total weight, lb
Private p_ca        As Double   ' Cushion pressure, lb/ft²
Private delta_p     As Double   ' Pressure differential, lb/ft²
Private Per_ag      As Double   ' Air gap perimeter, ft
Private A_ag        As Double   ' Air gap area, ft²
Private v_out       As Double   ' Air exit velocity, ft/s
Private v_in        As Double   ' Air inlet velocity, ft/s
Private V_dot       As Double   ' Volumetric flow rate, ft³/s
Private m_dot       As Double   ' Mass flow rate, slug/s
Private P_f         As Double   ' Fan power, ft·lb/s (internal)
Private p_ca_psi    As Double   ' Cushion pressure, psia (display)
Private delta_p_psi As Double   ' Pressure differential, psi (display)

' --- Unit conversion constants (US Customary to SI) ---
' Source: NIST SP811 Appendix B and derived values
' Multiply US value by factor to get SI value
' Divide SI value by factor to get US value
Public Const FT_TO_M        As Double = 0.3048              ' feet to meters (exact)
Public Const FT2_TO_M2      As Double = 0.09290304          ' ft² to m² (exact, derived)
Public Const FT3_TO_M3      As Double = 0.028316847         ' ft³ to m³ (derived)
Public Const SLUG_TO_KG     As Double = 14.593903           ' slugs to kilograms
Public Const SLUGFT3_TO_KGM3 As Double = 515.3788           ' slug/ft³ to kg/m³ (derived)
Public Const LT_TO_T        As Double = 1.0160469088        ' long tons to metric tonnes
Public Const LBFT2_TO_PA    As Double = 47.88026            ' lb/ft² to pascals (derived)
Public Const PSI_TO_PA      As Double = 6894.757            ' psi to pascals
Public Const FPS_TO_MPS     As Double = 0.3048              ' ft/s to m/s (exact, same as FT_TO_M)
Public Const FT3S_TO_M3S    As Double = 0.028316847         ' ft³/s to m³/s (same as FT3_TO_M3)
Public Const HP_TO_W        As Double = 745.699872          ' mechanical horsepower to watts


' ============================================================
' SUB: RunCalculations
'
' Master calculation routine. Called every time the user
' changes a value on the form. Reads form inputs, applies
' the correct equation rearrangement for each solve-for
' selection, and writes all results back to the form.
' ============================================================
Public Sub RunCalculations()

    ' Guard: do nothing if the form isn't loaded yet.
    ' This prevents errors during form initialization.
    If Not IsFormLoaded("frmHovercraft") Then Exit Sub

    

    ' --- Solve-for flags (one per bidirectional group) ---
    ' Each Boolean is True when that variable is the solve-for target.
    Dim solveWtot   As Boolean
    Dim solveWls    As Boolean
    Dim solveWdw    As Boolean
    Dim solveAc     As Boolean
    Dim solveLc     As Boolean
    Dim solveBc     As Boolean
    Dim solveDc     As Boolean
    Dim solveDsk    As Boolean
    Dim solveDag    As Boolean
    Dim solveDeltaP As Boolean
    Dim solvePca    As Boolean
    Dim solvePf     As Boolean
    Dim solveEta    As Boolean

    ' --- Program-wide mode flags ---
    ' useGauge is read from the form each time as it does not
    ' affect unit conversion -- only pressure calculation mode.
    ' isUS is a module-level variable and is NOT read from the
    ' form here -- it is set by the radio button click events
    ' and is the single source of truth for the unit system.
    Dim useGauge As Boolean  ' True = gauge pressure mode
    useGauge = frmHovercraft.optGauge.value

    ' ----------------------------------------------------------
    ' STEP 2: Read solve-for radio button states.
    ' .Value = True means that radio button is selected.
    ' ----------------------------------------------------------
    solveWtot = frmHovercraft.optSolveWtot.value
    solveWls = frmHovercraft.optSolveWls.value
    solveWdw = frmHovercraft.optSolveWdw.value
    solveAc = frmHovercraft.optSolveAc.value
    solveLc = frmHovercraft.optSolveLc.value
    solveBc = frmHovercraft.optSolveBc.value
    solveDc = frmHovercraft.optSolveDc.value
    solveDsk = frmHovercraft.optSolveDsk.value
    solveDag = frmHovercraft.optSolveDag.value
    solveDeltaP = frmHovercraft.optSolveDeltaP.value
    solvePca = frmHovercraft.optSolvePca.value
    solvePf = frmHovercraft.optSolvePf.value
    solveEta = frmHovercraft.optSolveEta.value

    ' ----------------------------------------------------------
    ' STEP 3: Read input values from the form text boxes.
    ' We skip reading a field if it is the solve-for target --
    ' its value will be overwritten by calculation anyway.
    ' SafeDouble() safely converts text to a number (returns 0
    ' for empty or non-numeric input -- see function below).
    ' ----------------------------------------------------------

    ' Atmospheric inputs always read (no solve-for on these).
    ' If SI is currently displayed, convert back to US base units
    ' before calculating. isUS = False means SI is displayed.
    p_a_psi = IIf(isUS, SafeDouble(frmHovercraft.txtPa.Text), _
              ToUS(SafeDouble(frmHovercraft.txtPa.Text), "pressure"))
    rho = IIf(isUS, SafeDouble(frmHovercraft.txtRho.Text), _
              ToUS(SafeDouble(frmHovercraft.txtRho.Text), "density"))
    p_a = p_a_psi * IN2_PER_FT2            ' convert psia to lb/ft²

    ' Flow coefficients
    ' C_d is active and user-adjustable, default value is 1.0
    ' C_v, C_c are reserved and inactive, hard-coded to default value of 1.0
    C_d = SafeDouble(frmHovercraft.txtCd.Text)
    C_v = 1#
    C_c = 1#

    ' Weight group
    ' Convert from display units back to US base units if in SI mode
    If Not solveWls Then
        W_ls_LT = IIf(isUS, SafeDouble(frmHovercraft.txtWls.Text), _
                  ToUS(SafeDouble(frmHovercraft.txtWls.Text), "weight"))
    End If
    If Not solveWdw Then
        W_dw_LT = IIf(isUS, SafeDouble(frmHovercraft.txtWdw.Text), _
                  ToUS(SafeDouble(frmHovercraft.txtWdw.Text), "weight"))
    End If
    If Not solveWtot Then
        W_tot_LT = IIf(isUS, SafeDouble(frmHovercraft.txtWtot.Text), _
                   ToUS(SafeDouble(frmHovercraft.txtWtot.Text), "weight"))
    End If
    W_ls = W_ls_LT * LB_PER_LT
    W_dw = W_dw_LT * LB_PER_LT
    W_tot = W_tot_LT * LB_PER_LT

    ' Cushion area group
    If Not solveLc Then
        L_c = IIf(isUS, SafeDouble(frmHovercraft.txtLc.Text), _
              ToUS(SafeDouble(frmHovercraft.txtLc.Text), "length"))
    End If
    If Not solveBc Then
        B_c = IIf(isUS, SafeDouble(frmHovercraft.txtBc.Text), _
              ToUS(SafeDouble(frmHovercraft.txtBc.Text), "length"))
    End If
    If Not solveAc Then
        A_c = IIf(isUS, SafeDouble(frmHovercraft.txtAc.Text), _
              ToUS(SafeDouble(frmHovercraft.txtAc.Text), "area"))
    End If

    ' Depth group
    ' theta is reserved and inactive, hard-coded to default value of 0
    If Not solveDsk Then
        D_sk = IIf(isUS, SafeDouble(frmHovercraft.txtDsk.Text), _
               ToUS(SafeDouble(frmHovercraft.txtDsk.Text), "length"))
    End If
    If Not solveDag Then
        D_ag = IIf(isUS, SafeDouble(frmHovercraft.txtDag.Text), _
               ToUS(SafeDouble(frmHovercraft.txtDag.Text), "length"))
    End If
    If Not solveDc Then
        D_c = IIf(isUS, SafeDouble(frmHovercraft.txtDc.Text), _
              ToUS(SafeDouble(frmHovercraft.txtDc.Text), "length"))
    End If
    theta = 0#

    ' Fan group
    If Not solveEta Then
        eta = SafeDouble(frmHovercraft.txtEta.Text)    ' dimensionless, no conversion
    End If
    If Not solvePf Then
        P_f_hp = IIf(isUS, SafeDouble(frmHovercraft.txtPf.Text), _
                 ToUS(SafeDouble(frmHovercraft.txtPf.Text), "power"))
    End If

    ' ----------------------------------------------------------
    ' STEP 4: Solve each bidirectional group.
    ' The If/ElseIf structure picks the correct algebraic form
    ' of the equation based on the active solve-for selection.
    ' ----------------------------------------------------------

    ' --- Weight group ---
    ' Base equation:    W_tot = W_ls + W_dw
    ' Solve for W_ls:   W_ls  = W_tot - W_dw
    ' Solve for W_dw:   W_dw  = W_tot - W_ls
    If solveWtot Then
        W_tot = W_ls + W_dw
        W_tot_LT = W_tot / LB_PER_LT
    ElseIf solveWls Then
        W_ls = W_tot - W_dw
        W_ls_LT = W_ls / LB_PER_LT
    ElseIf solveWdw Then
        W_dw = W_tot - W_ls
        W_dw_LT = W_dw / LB_PER_LT
    End If

    ' --- Cushion area group ---
    ' Base equation:   A_c = L_c · B_c
    ' Solve for L_c:   L_c = A_c / B_c
    ' Solve for B_c:   B_c = A_c / L_c
    ' Guards prevent division by zero.
    If solveAc Then
        A_c = L_c * B_c
    ElseIf solveLc Then
        If B_c <> 0 Then L_c = A_c / B_c
    ElseIf solveBc Then
        If L_c <> 0 Then B_c = A_c / L_c
    End If

    ' --- Depth group ---
    ' Base equation:    D_c  = D_sk + D_ag
    ' Solve for D_sk:   D_sk = D_c - D_ag
    ' Solve for D_ag:   D_ag = D_c - D_sk
    If solveDc Then
        D_c = D_sk + D_ag
    ElseIf solveDsk Then
        D_sk = D_c - D_ag
    ElseIf solveDag Then
        D_ag = D_c - D_sk
    End If

    ' --- Derived geometry (always calculated, no solve-for) ---
    Per_ag = 2 * (L_c + B_c)    ' Air gap perimeter, ft
    A_ag = Per_ag * D_ag         ' Air gap area, ft²

    ' --- Cushion pressure & Delta P ---
    ' In gauge mode, cushion pressure is simply the pressure required
    ' to support the vessel weight over the cushion area.
    ' In absolute mode, cushion pressure is the gauge pressure PLUS
    ' ambient pressure, because absolute pressure is measured from
    ' zero rather than from ambient.
    ' Guard against division by zero if A_c hasn't been entered yet.
    ' --- Delta P ---
    ' Absolute mode:  delta_p = p_ca - p_a
    ' Gauge mode:     delta_p = p_ca
    '   (gauge pressure is already expressed relative to ambient,
    '    so ambient is implicitly zero in gauge mode)
    If A_c <> 0 Then
        If useGauge Then
            p_ca = W_tot / A_c              ' lb/ft² gauge
            delta_p = p_ca                  ' gauge delta p = cushion pressure
        Else
            p_ca = (W_tot / A_c) + p_a      ' lb/ft² absolute
            delta_p = p_ca - p_a            ' absolute delta p = cushion - ambient
        End If
    Else
        p_ca = 0
        delta_p = 0
    End If

    ' --- Flow calculations ---
    '   v_out = C_d · sqrt(2 · delta_p / rho)
    '   Torricelli/Bernoulli jet velocity equation.
    '   Requires delta_p >= 0 and rho > 0 to be physically valid.
    If rho > 0 And delta_p >= 0 Then
        v_out = C_d * Sqr(2 * delta_p / rho)
    Else
        v_out = 0
    End If

    V_dot = A_ag * v_out         ' Volumetric flow rate, ft³/s
    m_dot = V_dot * rho          ' Mass flow rate, slug/s

    ' v_in reserved -- will use V_dot / A_fan when fan geometry is active
    v_in = 0

    ' --- Fan power group ---
    ' Base equation:   P_f = (delta_p · V_dot) / eta   [ft·lb/s]
    ' Solve for eta:   eta = (delta_p · V_dot) / P_f
    ' Convert between hp and ft·lb/s using FTLBS_PER_HP.
    If solvePf Then
        If eta > 0 Then
            P_f = (delta_p * V_dot) / eta
            P_f_hp = P_f / FTLBS_PER_HP
        Else
            P_f_hp = 0
        End If
    ElseIf solveEta Then
        P_f = P_f_hp * FTLBS_PER_HP
        If P_f > 0 Then
            eta = (delta_p * V_dot) / P_f
        Else
            eta = 0
        End If
    End If

    ' ----------------------------------------------------------
    ' STEP 5: Populate dictionaries and update the display.
    ' All calculated values are now stored in both US and SI
    ' dictionaries. DisplayFromDictionary then writes whichever
    ' set matches the current unit system to the form.
    ' ----------------------------------------------------------
    Call PopulateDictionaries
    Call DisplayFromDictionary
    Call UpdateFieldColors

End Sub


' ============================================================
' SUB: UpdateFieldColors
'
' Sets the BackColor of every text box to visually show
' whether it is a user-input field (white) or a calculated
' field (light grey) or a reserved field (darker grey).
' Called after every calculation run.
' ============================================================
Public Sub UpdateFieldColors()

    Const INPUT_COLOR   As Long = &HFFFFFF  ' White -- user input
    Const CALC_COLOR    As Long = &HF2F3F4  ' Light grey -- calculated
    Const LOCKED_COLOR  As Long = &HE0E0E0  ' Medium grey -- reserved

    With frmHovercraft

        ' Weight group: the solve-for field gets CALC_COLOR
        .txtWls.BackColor = IIf(.optSolveWls.value, CALC_COLOR, INPUT_COLOR)
        .txtWdw.BackColor = IIf(.optSolveWdw.value, CALC_COLOR, INPUT_COLOR)
        .txtWtot.BackColor = IIf(.optSolveWtot.value, CALC_COLOR, INPUT_COLOR)

        ' Cushion area group
        .txtLc.BackColor = IIf(.optSolveLc.value, CALC_COLOR, INPUT_COLOR)
        .txtBc.BackColor = IIf(.optSolveBc.value, CALC_COLOR, INPUT_COLOR)
        .txtAc.BackColor = IIf(.optSolveAc.value, CALC_COLOR, INPUT_COLOR)

        ' Depth group
        .txtDsk.BackColor = IIf(.optSolveDsk.value, CALC_COLOR, INPUT_COLOR)
        .txtDag.BackColor = IIf(.optSolveDag.value, CALC_COLOR, INPUT_COLOR)
        .txtDc.BackColor = IIf(.optSolveDc.value, CALC_COLOR, INPUT_COLOR)
        .txtTheta.BackColor = LOCKED_COLOR

        ' Always-calculated fields (no solve-for option on these)
        .txtPerag.BackColor = CALC_COLOR
        .txtAag.BackColor = CALC_COLOR
        .txtVin.BackColor = CALC_COLOR
        .txtVout.BackColor = CALC_COLOR
        .txtVdot.BackColor = CALC_COLOR
        .txtMdot.BackColor = CALC_COLOR

        ' Pressure group
        .txtPca.BackColor = IIf(.optSolvePca.value, CALC_COLOR, INPUT_COLOR)
        .txtDeltaP.BackColor = IIf(.optSolveDeltaP.value, CALC_COLOR, INPUT_COLOR)
        .txtCv.BackColor = LOCKED_COLOR
        .txtCc.BackColor = LOCKED_COLOR
        .txtCd.BackColor = INPUT_COLOR

        ' Fan power group
        .txtEta.BackColor = IIf(.optSolveEta.value, CALC_COLOR, INPUT_COLOR)
        .txtPf.BackColor = IIf(.optSolvePf.value, CALC_COLOR, INPUT_COLOR)

        ' Reserved fan geometry fields always locked
        .txtRfan.BackColor = LOCKED_COLOR
        .txtAfan.BackColor = LOCKED_COLOR
        .txtNfan.BackColor = LOCKED_COLOR

    End With

End Sub

' ============================================================
' SUB: PopulateDictionaries
'
' Called at the end of every RunCalculations cycle.
' Stores all current US Customary values in dictUSValues and
' all converted SI values in dictSIValues.
' Also stores unit label strings in dictUSUnits and dictSIUnits.
'
' Both dictionaries are always fully populated after every
' calculation so switching unit systems requires no recalculation
' -- DisplayFromDictionary simply reads from the correct one.
' ============================================================
Public Sub PopulateDictionaries()

    ' Initialize all four dictionaries fresh each call.
    ' This ensures stale values from previous calculations
    ' are never accidentally displayed.
    Set dictUSValues = New Scripting.Dictionary
    Set dictSIValues = New Scripting.Dictionary
    Set dictUSUnits = New Scripting.Dictionary
    Set dictSIUnits = New Scripting.Dictionary

    ' Convert lb/ft² pressure values to psia for display
    p_ca_psi = p_ca / IN2_PER_FT2
    delta_p_psi = delta_p / IN2_PER_FT2

    ' ----------------------------------------------------------
    ' US Customary values (always in base US units)
    ' Key = TextBox name string, Value = US display value
    ' ----------------------------------------------------------
    dictUSValues("txtPa") = p_a_psi
    dictUSValues("txtRho") = rho
    dictUSValues("txtWls") = W_ls_LT
    dictUSValues("txtWdw") = W_dw_LT
    dictUSValues("txtWtot") = W_tot_LT
    dictUSValues("txtLCG") = SafeDouble(frmHovercraft.txtLCG.Text)
    dictUSValues("txtTCG") = SafeDouble(frmHovercraft.txtTCG.Text)
    dictUSValues("txtVCGf") = SafeDouble(frmHovercraft.txtVCGf.Text)
    dictUSValues("txtVCGc") = SafeDouble(frmHovercraft.txtVCGc.Text)
    dictUSValues("txtLc") = L_c
    dictUSValues("txtBc") = B_c
    dictUSValues("txtAc") = A_c
    dictUSValues("txtDsk") = D_sk
    dictUSValues("txtDag") = D_ag
    dictUSValues("txtDc") = D_c
    dictUSValues("txtPerag") = Per_ag
    dictUSValues("txtAag") = A_ag
    dictUSValues("txtPca") = p_ca_psi
    dictUSValues("txtDeltaP") = delta_p_psi
    dictUSValues("txtVin") = v_in
    dictUSValues("txtVout") = v_out
    dictUSValues("txtVdot") = V_dot
    dictUSValues("txtMdot") = m_dot
    dictUSValues("txtEta") = eta
    dictUSValues("txtCd") = C_d
    dictUSValues("txtPf") = P_f_hp

    ' ----------------------------------------------------------
    ' SI values (converted from US base units)
    ' Key = TextBox name string, Value = SI display value
    ' ----------------------------------------------------------
    dictSIValues("txtPa") = ToSI(p_a_psi, "pressure")
    dictSIValues("txtRho") = ToSI(rho, "density")
    dictSIValues("txtWls") = ToSI(W_ls_LT, "weight")
    dictSIValues("txtWdw") = ToSI(W_dw_LT, "weight")
    dictSIValues("txtWtot") = ToSI(W_tot_LT, "weight")
    dictSIValues("txtLCG") = ToSI(SafeDouble(frmHovercraft.txtLCG.Text), "length")
    dictSIValues("txtTCG") = ToSI(SafeDouble(frmHovercraft.txtTCG.Text), "length")
    dictSIValues("txtVCGf") = ToSI(SafeDouble(frmHovercraft.txtVCGf.Text), "length")
    dictSIValues("txtVCGc") = ToSI(SafeDouble(frmHovercraft.txtVCGc.Text), "length")
    dictSIValues("txtLc") = ToSI(L_c, "length")
    dictSIValues("txtBc") = ToSI(B_c, "length")
    dictSIValues("txtAc") = ToSI(A_c, "area")
    dictSIValues("txtDsk") = ToSI(D_sk, "length")
    dictSIValues("txtDag") = ToSI(D_ag, "length")
    dictSIValues("txtDc") = ToSI(D_c, "length")
    dictSIValues("txtPerag") = ToSI(Per_ag, "length")
    dictSIValues("txtAag") = ToSI(A_ag, "area")
    dictSIValues("txtPca") = ToSI(p_ca_psi, "pressure")
    dictSIValues("txtDeltaP") = ToSI(delta_p_psi, "pressure")
    dictSIValues("txtVin") = ToSI(v_in, "velocity")
    dictSIValues("txtVout") = ToSI(v_out, "velocity")
    dictSIValues("txtVdot") = ToSI(V_dot, "volflow")
    dictSIValues("txtMdot") = ToSI(m_dot, "massflow")
    dictSIValues("txtEta") = eta            ' dimensionless -- no conversion
    dictSIValues("txtCd") = C_d             ' dimensionless -- no conversion
    dictSIValues("txtPf") = ToSI(P_f_hp, "power")

    ' ----------------------------------------------------------
    ' US unit label strings
    ' Key = Label control name, Value = US unit string
    ' ----------------------------------------------------------
    dictUSUnits("lblPaUnit") = "psia"
    dictUSUnits("lblRhoUnit") = "slug/ft³"
    dictUSUnits("lblWlsUnit") = "LT"
    dictUSUnits("lblWdwUnit") = "LT"
    dictUSUnits("lblWtotUnit") = "LT"
    dictUSUnits("lblLCGUnit") = "ft"
    dictUSUnits("lblTCGUnit") = "ft"
    dictUSUnits("lblVCGfUnit") = "ft"
    dictUSUnits("lblVCGcUnit") = "ft"
    dictUSUnits("lblLcUnit") = "ft"
    dictUSUnits("lblBcUnit") = "ft"
    dictUSUnits("lblAcUnit") = "ft²"
    dictUSUnits("lblDskUnit") = "ft"
    dictUSUnits("lblDagUnit") = "ft"
    dictUSUnits("lblDcUnit") = "ft"
    dictUSUnits("lblPeragUnit") = "ft"
    dictUSUnits("lblAagUnit") = "ft²"
    dictUSUnits("lblPcaUnit") = "psia"
    dictUSUnits("lblDeltaPUnit") = "psi"
    dictUSUnits("lblVinUnit") = "ft/s"
    dictUSUnits("lblVoutUnit") = "ft/s"
    dictUSUnits("lblVdotUnit") = "ft³/s"
    dictUSUnits("lblMdotUnit") = "slug/s"
    dictUSUnits("lblEtaUnit") = "[ - ]"
    dictUSUnits("lblCdUnit") = "[ - ]"
    dictUSUnits("lblPfUnit") = "hp"
    dictUSUnits("lblRfanUnit") = "ft"
    dictUSUnits("lblAfanUnit") = "ft²"
    dictUSUnits("lblNfanUnit") = "rpm"

    ' ----------------------------------------------------------
    ' SI unit label strings
    ' Key = Label control name, Value = SI unit string
    ' ----------------------------------------------------------
    dictSIUnits("lblPaUnit") = "Pa"
    dictSIUnits("lblRhoUnit") = "kg/m³"
    dictSIUnits("lblWlsUnit") = "t"
    dictSIUnits("lblWdwUnit") = "t"
    dictSIUnits("lblWtotUnit") = "t"
    dictSIUnits("lblLCGUnit") = "m"
    dictSIUnits("lblTCGUnit") = "m"
    dictSIUnits("lblVCGfUnit") = "m"
    dictSIUnits("lblVCGcUnit") = "m"
    dictSIUnits("lblLcUnit") = "m"
    dictSIUnits("lblBcUnit") = "m"
    dictSIUnits("lblAcUnit") = "m²"
    dictSIUnits("lblDskUnit") = "m"
    dictSIUnits("lblDagUnit") = "m"
    dictSIUnits("lblDcUnit") = "m"
    dictSIUnits("lblPeragUnit") = "m"
    dictSIUnits("lblAagUnit") = "m²"
    dictSIUnits("lblPcaUnit") = "Pa"
    dictSIUnits("lblDeltaPUnit") = "Pa"
    dictSIUnits("lblVinUnit") = "m/s"
    dictSIUnits("lblVoutUnit") = "m/s"
    dictSIUnits("lblVdotUnit") = "m³/s"
    dictSIUnits("lblMdotUnit") = "kg/s"
    dictSIUnits("lblEtaUnit") = "[ - ]"
    dictSIUnits("lblCdUnit") = "[ - ]"
    dictSIUnits("lblPfUnit") = "W"
    dictSIUnits("lblRfanUnit") = "m"
    dictSIUnits("lblAfanUnit") = "m²"
    dictSIUnits("lblNfanUnit") = "rpm"

End Sub


' ============================================================
' SUB: DisplayFromDictionary
'
' Reads from dictUSValues/dictUSUnits or dictSIValues/dictSIUnits
' depending on the current value of isUS, and writes all values
' and unit labels to the form.
'
' Called after every PopulateDictionaries call and also directly
' from the unit system radio button click events -- in the latter
' case no recalculation happens, just a redisplay from the
' already-populated dictionaries.
' ============================================================
Public Sub DisplayFromDictionary()

    ' Select which dictionaries to read from
    Dim vals  As Scripting.Dictionary
    Dim units As Scripting.Dictionary

    If isUS Then
        Set vals = dictUSValues
        Set units = dictUSUnits
    Else
        Set vals = dictSIValues
        Set units = dictSIUnits
    End If

    ' Prevent change events from firing while writing to form
    frmHovercraft.m_calculating = True

    ' ----------------------------------------------------------
    ' Write values to TextBoxes
    ' Decimal places per field match the original WriteField calls
    ' ----------------------------------------------------------
    WriteField frmHovercraft.txtPa, vals("txtPa"), 4
    WriteField frmHovercraft.txtRho, vals("txtRho"), 5
    WriteField frmHovercraft.txtWls, vals("txtWls"), 4
    WriteField frmHovercraft.txtWdw, vals("txtWdw"), 4
    WriteField frmHovercraft.txtWtot, vals("txtWtot"), 4
    WriteField frmHovercraft.txtLCG, vals("txtLCG"), 3
    WriteField frmHovercraft.txtTCG, vals("txtTCG"), 3
    WriteField frmHovercraft.txtVCGf, vals("txtVCGf"), 3
    WriteField frmHovercraft.txtVCGc, vals("txtVCGc"), 3
    WriteField frmHovercraft.txtLc, vals("txtLc"), 3
    WriteField frmHovercraft.txtBc, vals("txtBc"), 3
    WriteField frmHovercraft.txtAc, vals("txtAc"), 3
    WriteField frmHovercraft.txtDsk, vals("txtDsk"), 3
    WriteField frmHovercraft.txtDag, vals("txtDag"), 4
    WriteField frmHovercraft.txtDc, vals("txtDc"), 3
    WriteField frmHovercraft.txtPerag, vals("txtPerag"), 3
    WriteField frmHovercraft.txtAag, vals("txtAag"), 4
    WriteField frmHovercraft.txtPca, vals("txtPca"), 4
    WriteField frmHovercraft.txtDeltaP, vals("txtDeltaP"), 6
    WriteField frmHovercraft.txtVin, vals("txtVin"), 3
    WriteField frmHovercraft.txtVout, vals("txtVout"), 3
    WriteField frmHovercraft.txtVdot, vals("txtVdot"), 3
    WriteField frmHovercraft.txtMdot, vals("txtMdot"), 5
    WriteField frmHovercraft.txtEta, vals("txtEta"), 4
    WriteField frmHovercraft.txtCd, vals("txtCd"), 4
    WriteField frmHovercraft.txtPf, vals("txtPf"), 3

    ' ----------------------------------------------------------
    ' Write unit label captions
    ' ----------------------------------------------------------
    With frmHovercraft
        .lblPaUnit.Caption = units("lblPaUnit")
        .lblRhoUnit.Caption = units("lblRhoUnit")
        .lblWlsUnit.Caption = units("lblWlsUnit")
        .lblWdwUnit.Caption = units("lblWdwUnit")
        .lblWtotUnit.Caption = units("lblWtotUnit")
        .lblLCGUnit.Caption = units("lblLCGUnit")
        .lblTCGUnit.Caption = units("lblTCGUnit")
        .lblVCGfUnit.Caption = units("lblVCGfUnit")
        .lblVCGcUnit.Caption = units("lblVCGcUnit")
        .lblLcUnit.Caption = units("lblLcUnit")
        .lblBcUnit.Caption = units("lblBcUnit")
        .lblAcUnit.Caption = units("lblAcUnit")
        .lblDskUnit.Caption = units("lblDskUnit")
        .lblDagUnit.Caption = units("lblDagUnit")
        .lblDcUnit.Caption = units("lblDcUnit")
        .lblPeragUnit.Caption = units("lblPeragUnit")
        .lblAagUnit.Caption = units("lblAagUnit")
        .lblPcaUnit.Caption = units("lblPcaUnit")
        .lblDeltaPUnit.Caption = units("lblDeltaPUnit")
        .lblVinUnit.Caption = units("lblVinUnit")
        .lblVoutUnit.Caption = units("lblVoutUnit")
        .lblVdotUnit.Caption = units("lblVdotUnit")
        .lblMdotUnit.Caption = units("lblMdotUnit")
        .lblEtaUnit.Caption = units("lblEtaUnit")
        .lblCdUnit.Caption = units("lblCdUnit")
        .lblPfUnit.Caption = units("lblPfUnit")
        .lblRfanUnit.Caption = units("lblRfanUnit")
        .lblAfanUnit.Caption = units("lblAfanUnit")
        .lblNfanUnit.Caption = units("lblNfanUnit")
    End With

    frmHovercraft.m_calculating = False

End Sub

' ============================================================
' SUB: RecordScenario
'
' Appends the current form values as a new row in the
' Scenarios sheet. Called by the Record Scenario button.
' ============================================================
Public Sub RecordScenario()

    Dim ws      As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets("Scenarios")

    ' Find the next empty row â€” data starts at row 4
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 4 Then nextRow = 4

    With frmHovercraft

        ws.Cells(nextRow, 1).value = nextRow - 3            ' Scenario number
        ws.Cells(nextRow, 2).value = Now()                  ' Timestamp
        ws.Cells(nextRow, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ws.Cells(nextRow, 3).value = ""                     ' Description blank for user
        ws.Cells(nextRow, 4).value = IIf(.optSI.value, "SI", "US Customary")
        ws.Cells(nextRow, 5).value = IIf(.optGauge.value, "Gauge", "Absolute")

        ws.Cells(nextRow, 6).value = SafeDouble(.txtWls.Text)
        ws.Cells(nextRow, 7).value = SafeDouble(.txtWdw.Text)
        ws.Cells(nextRow, 8).value = SafeDouble(.txtWtot.Text)
        ws.Cells(nextRow, 9).value = SafeDouble(.txtLCG.Text)
        ws.Cells(nextRow, 10).value = SafeDouble(.txtTCG.Text)
        ws.Cells(nextRow, 11).value = SafeDouble(.txtVCGf.Text)
        ws.Cells(nextRow, 12).value = SafeDouble(.txtVCGc.Text)

        ws.Cells(nextRow, 13).value = SafeDouble(.txtPa.Text)
        ws.Cells(nextRow, 14).value = SafeDouble(.txtRho.Text)

        ws.Cells(nextRow, 15).value = SafeDouble(.txtLc.Text)
        ws.Cells(nextRow, 16).value = SafeDouble(.txtBc.Text)
        ws.Cells(nextRow, 17).value = SafeDouble(.txtAc.Text)
        ws.Cells(nextRow, 18).value = SafeDouble(.txtDsk.Text)
        ws.Cells(nextRow, 19).value = SafeDouble(.txtDag.Text)
        ws.Cells(nextRow, 20).value = SafeDouble(.txtDc.Text)
        ws.Cells(nextRow, 21).value = SafeDouble(.txtPerag.Text)
        ws.Cells(nextRow, 22).value = SafeDouble(.txtAag.Text)

        ws.Cells(nextRow, 23).value = SafeDouble(.txtPca.Text)
        ws.Cells(nextRow, 24).value = SafeDouble(.txtDeltaP.Text)
        ws.Cells(nextRow, 25).value = SafeDouble(.txtVin.Text)
        ws.Cells(nextRow, 26).value = SafeDouble(.txtVout.Text)
        ws.Cells(nextRow, 27).value = SafeDouble(.txtVdot.Text)
        ws.Cells(nextRow, 28).value = SafeDouble(.txtMdot.Text)
        ws.Cells(nextRow, 29).value = SafeDouble(.txtCd.Text)

        ws.Cells(nextRow, 30).value = SafeDouble(.txtEta.Text)
        ws.Cells(nextRow, 31).value = SafeDouble(.txtPf.Text)

    End With

    With ws.Rows(nextRow)
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
    End With

    MsgBox "Scenario " & (nextRow - 3) & " recorded." & vbNewLine & _
           "Switch to the Scenarios sheet to view and add a description.", _
           vbInformation, "Scenario Recorded"

End Sub


' ============================================================
' SUB: ResetForm
' Restores all fields and radio buttons to default values.
' ============================================================
Public Sub ResetForm()

    If MsgBox("Reset all fields to default values?", _
              vbYesNo + vbQuestion, "Reset") = vbNo Then Exit Sub

    With frmHovercraft
        .m_calculating = True

        .txtPa.Text = "14.696"
        .txtRho.Text = "0.00237"
        .txtWls.Text = "0.000"
        .txtWdw.Text = "0.000"
        .txtWtot.Text = "0.000"
        .txtLCG.Text = "0.000"
        .txtTCG.Text = "0.000"
        .txtVCGf.Text = "0.000"
        .txtVCGc.Text = "0.000"
        .txtLc.Text = "0.000"
        .txtBc.Text = "0.000"
        .txtAc.Text = "0.000"
        .txtDsk.Text = "0.000"
        .txtDag.Text = "0.000"
        .txtDc.Text = "0.000"
        .txtCd.Text = "1.000"
        .txtEta.Text = "1.000"
        .txtPf.Text = "0.000"

        ' Restore default solve-for selections
        .optSolveWtot.value = True
        .optSolveAc.value = True
        .optSolveDc.value = True
        .optSolveDeltaP.value = True
        .optSolvePf.value = True

        ' Restore program-wide defaults
        .optUSCustomary.value = True
        .optAbsolute.value = True

        .m_calculating = False
    End With

    Call RunCalculations

End Sub


' ============================================================
' SUB: ClearForm
' Zeros all input values while preserving radio button state.
' ============================================================
Public Sub ClearForm()

    If MsgBox("Clear all input values?", _
              vbYesNo + vbQuestion, "Clear") = vbNo Then Exit Sub

    With frmHovercraft
        .m_calculating = True

        .txtWls.Text = "0.000"
        .txtWdw.Text = "0.000"
        .txtWtot.Text = "0.000"
        .txtLCG.Text = "0.000"
        .txtTCG.Text = "0.000"
        .txtVCGf.Text = "0.000"
        .txtVCGc.Text = "0.000"
        .txtLc.Text = "0.000"
        .txtBc.Text = "0.000"
        .txtAc.Text = "0.000"
        .txtDsk.Text = "0.000"
        .txtDag.Text = "0.000"
        .txtDc.Text = "0.000"
        .txtCd.Text = "1.000"
        .txtEta.Text = "1.000"
        .txtPf.Text = "0.000"

        .m_calculating = False
    End With

    Call RunCalculations

End Sub


' ============================================================
' FUNCTION: SafeDouble
' Converts a string to Double; returns 0 for blank/invalid.
' ============================================================
Public Function SafeDouble(ByVal s As String) As Double
    If IsNumeric(Trim(s)) Then
        SafeDouble = CDbl(Trim(s))
    Else
        SafeDouble = 0#
    End If
End Function


' ============================================================
' SUB: WriteField
' Writes a formatted number to a TextBox control.
' Parameters:
'   ctrl     - the TextBox to write to
'   value    - the numeric value
'   decimals - decimal places to display
' ============================================================
Public Sub WriteField(ByRef ctrl As MSForms.TextBox, _
                      ByVal value As Double, _
                      ByVal decimals As Integer)
    ctrl.Text = Format(value, "0." & String(decimals, "0"))
End Sub


' ============================================================
' FUNCTION: IsFormLoaded
' Returns True if the named UserForm is currently in memory.
' ============================================================
Public Function IsFormLoaded(ByVal formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

' ============================================================
' FUNCTION: ToSI
'
' Converts a value from US Customary to SI units.
' Returns the converted value, or the original value unchanged
' if the unit type is not recognized or conversion is not needed.
'
' Parameters:
'   value    - the numeric value in US Customary units
'   unitType - a string identifying what kind of unit this is
'              so the correct conversion factor can be applied
'
' Usage example:
'   siValue = ToSI(myLengthInFeet, "length")
' ============================================================
Public Function ToSI(ByVal value As Double, _
                     ByVal unitType As String) As Double

    ' Load conversion factors into a dictionary.
    ' Key   = unit type string (matches what callers pass in)
    ' Value = multiplication factor to convert US to SI
    Dim factors As New Scripting.Dictionary
    factors("length") = FT_TO_M
    factors("area") = FT2_TO_M2
    factors("volume") = FT3_TO_M3
    factors("density") = SLUGFT3_TO_KGM3
    factors("mass") = SLUG_TO_KG
    factors("massflow") = SLUG_TO_KG
    factors("weight") = LT_TO_T
    factors("pressure") = PSI_TO_PA
    factors("velocity") = FPS_TO_MPS
    factors("volflow") = FT3S_TO_M3S
    factors("power") = HP_TO_W

    ' If the unit type exists in the dictionary, apply the factor.
    ' If not, return the value unchanged -- dimensionless values
    ' like eta and Cd do not need conversion.
    If factors.Exists(unitType) Then
        ToSI = value * factors(unitType)
    Else
        ToSI = value
    End If

End Function


' ============================================================
' FUNCTION: ToUS
'
' Converts a value from SI back to US Customary units.
' This is the reverse of ToSI -- divides by the same factor
' rather than multiplying.
'
' Parameters:
'   value    - the numeric value in SI units
'   unitType - same unit type string as used in ToSI
'
' Usage example:
'   usValue = ToUS(myLengthInMeters, "length")
' ============================================================
Public Function ToUS(ByVal value As Double, _
                     ByVal unitType As String) As Double

    ' We reuse the same dictionary and factors as ToSI.
    ' To convert SI back to US we simply divide by the factor
    ' instead of multiplying -- the inverse operation.
    Dim factors As New Scripting.Dictionary
    factors("length") = FT_TO_M
    factors("area") = FT2_TO_M2
    factors("volume") = FT3_TO_M3
    factors("density") = SLUGFT3_TO_KGM3
    factors("mass") = SLUG_TO_KG
    factors("massflow") = SLUG_TO_KG
    factors("weight") = LT_TO_T
    factors("pressure") = PSI_TO_PA
    factors("velocity") = FPS_TO_MPS
    factors("volflow") = FT3S_TO_M3S
    factors("power") = HP_TO_W

    If factors.Exists(unitType) Then
        ToUS = value / factors(unitType)
    Else
        ToUS = value
    End If

End Function

