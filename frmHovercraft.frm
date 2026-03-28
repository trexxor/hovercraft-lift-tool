VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHovercraft 
   Caption         =   "Hovercraft Performance Calculator"
   ClientHeight    =   13995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   OleObjectBlob   =   "frmHovercraft.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHovercraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' USERFORM MODULE: frmHovercraft
'
' PURPOSE:
'   This code belongs inside the UserForm named frmHovercraft.
'   It handles all user interactions ON the form — button
'   clicks, text changes, radio button selections — and
'   calls the calculation engine in HovercraftCalc.bas.
'
' HOW TO INSTALL THIS CODE:
'   1. Press Alt+F11 to open the VBA Editor
'   2. In the Project panel, right-click your project name
'   3. Choose Insert > UserForm
'   4. Rename the form to "frmHovercraft" in the Properties panel
'   5. Double-click the form to open its code window
'   6. Paste this entire file into that code window
'
' NOTE ON BUILDING THE FORM VISUALLY:
'   The controls (text boxes, labels, radio buttons, buttons)
'   must be drawn manually on the form canvas in the VBA Editor.
'   See FORM_SETUP_GUIDE.txt for the complete list of control
'   names, types, and layout instructions.
'
' CONTROL NAMING CONVENTION:
'   txt  = TextBox      (e.g. txtWls, txtPa)
'   opt  = OptionButton (radio button, e.g. optSolveWtot)
'   lbl  = Label        (e.g. lblWls)
'   cmd  = CommandButton (e.g. cmdRecord)
'   frm  = Frame        (grouping container, e.g. frmMass)
' ============================================================

' ============================================================
' WINDOWS API DECLARATIONS — Minimize/Restore button support
'
' These declarations tell VBA where to find the Windows
' functions that control the form's title bar buttons.
'
' The #If Win64 block handles both 32-bit and 64-bit Excel
' automatically — VBA picks the correct declaration at
' compile time based on which version is running.
'
' GetWindowLongPtr — reads the current window style settings
' SetWindowLongPtr — writes new window style settings
' FindWindowA      — finds the form's window handle by name
' DrawMenuBar      — redraws the title bar after changes
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLongPtr _
        Lib "user32.dll" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long) As LongPtr

    Private Declare PtrSafe Function SetWindowLongPtr _
        Lib "user32.dll" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) As LongPtr

    Private Declare PtrSafe Function FindWindowA _
        Lib "user32.dll" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar _
        Lib "user32.dll" ( _
        ByVal hwnd As LongPtr) As Long
#End If

' Window style constants
Private Const GWL_STYLE     As Long = -16       ' Style offset
Private Const WS_MINIMIZEBOX As Long = &H20000  ' Minimize button flag

' ----------------------------------------------------------
' Module-level collection that holds all clsTextBox instances.
'
' This must be declared at the module level, not inside a Sub.
' If it were declared inside a Sub it would be destroyed when
' that Sub finished, taking all the event connections with it
' and leaving the TextBoxes unresponsive.
'
' A Collection is a VBA built-in object that holds a group
' of items -- similar to a list. We use it here purely as a
' container to keep the instances alive in memory for as long
' as the form is open.
' ----------------------------------------------------------
Private m_TextBoxHandlers As Collection


' -- Module-level flag ------
' m_calculating is True while RunCalculations() is writing values
' back to the form. This prevents the Change events on text boxes
' from firing during that write-back and triggering a loop.
' "Public" makes it readable from the HovercraftCalc module.
Public m_calculating As Boolean


' ============================================================
' EVENT: UserForm_Initialize
'
' This event fires automatically when the form is first loaded
' into memory (before it becomes visible on screen).
' We use it to set all default values and radio button states.
' ============================================================
Private Sub UserForm_Initialize()
    
    ' ----------------------------------------------------------
    ' Initialize the collection that will hold all clsTextBox
    ' instances. This must be done before the loop below that
    ' populates it -- you cannot add items to a collection that
    ' has not been created yet.
    ' ----------------------------------------------------------
    Set m_TextBoxHandlers = New Collection
    
    ' Set the flag so no calculations fire during initialization
    m_calculating = True
    
    ' Initialize the unit system state and dictionaries.
    ' isUS = True means US Customary is the default display unit system.
    ' Dictionaries must be instantiated before RunCalculations is called
    ' for the first time at the bottom of this sub.
    isUS = True
    Set dictUSValues = New Scripting.Dictionary
    Set dictSIValues = New Scripting.Dictionary
    Set dictUSUnits = New Scripting.Dictionary
    Set dictSIUnits = New Scripting.Dictionary

    ' --- Default atmospheric values (sea level, standard conditions) ---
    txtPa.Text = "14.696"     ' psia — sea level standard
    txtRho.Text = "0.00237"   ' slug/ft³ — sea level, 59°F standard

    ' --- Default weight values ---
    txtWls.Text = "0.000"
    txtWdw.Text = "0.000"
    txtWtot.Text = "0.000"

    ' --- Default center of gravity values ---
    txtLCG.Text = "0.000"
    txtTCG.Text = "0.000"
    txtVCGf.Text = "0.000"
    txtVCGc.Text = "0.000"

    ' --- Default cushion geometry values ---
    txtLc.Text = "0.000"
    txtBc.Text = "0.000"
    txtAc.Text = "0.000"
    txtDsk.Text = "0.000"
    txtDag.Text = "0.000"
    txtDc.Text = "0.000"
    txtTheta.Text = "—"
    txtTheta.Enabled = False    ' Reserved awaiting equations for use
    
    ' --- Default flow coefficients values ---
    txtCv.Text = "—"
    txtCc.Text = "—"
    txtCd.Text = "1.000"
    txtCv.Enabled = False
    txtCc.Enabled = False
    
    ' --- Default fan values ---
    txtEta.Text = "1.000"
    txtPf.Text = "0.000"

    ' --- Reserved fan geometry fields display dashes ---
    txtRfan.Text = "—"
    txtAfan.Text = "—"
    txtNfan.Text = "—"
    txtRfan.Enabled = False
    txtAfan.Enabled = False
    txtNfan.Enabled = False

    ' --- Default solve-for radio button selections ---
    ' These match the agreed defaults from the design session.
    optSolveWtot.value = True     ' Weight group: solve for W_tot
    optSolveAc.value = True       ' Cushion area: solve for A_c
    optSolveDc.value = True       ' Depth group:  solve for D_c
    optSolveDeltaP.value = True   ' Pressure:     solve for delta_p
    optSolvePf.value = True       ' Fan power:    solve for P_f

    ' --- Default program-wide settings ---
    optUSCustomary.value = True   ' Units: US Customary
    optAbsolute.value = True      ' Pressure: Absolute
    
    ' ----------------------------------------------------------
    ' Create one clsTextBox instance for each enabled TextBox
    ' on the form and add it to the collection.
    '
    ' "Dim handler As clsTextBox" declares a variable that can
    ' hold one instance of our class. It is reused each time
    ' through the loop -- a new instance is created each
    ' iteration so reusing the variable causes no conflict.
    '
    ' "Dim ctrl As Control" declares a variable that temporarily
    ' holds each control as the loop cycles through Me.Controls.
    ' "Me" refers to this form. "Controls" is the form's built-in
    ' collection of every control placed on it.
    '
    ' "TypeName(ctrl) = TextBox" checks whether the current
    ' control is a TextBox. "ctrl.Enabled = True" excludes any
    ' TextBox that is currently disabled -- reserved fields like
    ' txtRfan will be excluded until they are activated.
    '
    ' "Set handler = New clsTextBox" creates a brand new instance
    ' from the clsTextBox blueprint and assigns it to handler.
    '
    ' "Set handler.TextBox = ctrl" connects the instance to the
    ' current TextBox -- this is the moment the WithEvents
    ' connection is made and the instance starts listening.
    '
    ' "m_TextBoxHandlers.Add handler" puts the instance into the
    ' collection to keep it alive in memory. Without this line
    ' the instance would be destroyed when the loop moves to the
    ' next iteration and the event connection would be lost.
    ' ----------------------------------------------------------
    Dim handler As clsTextBox
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" And ctrl.Enabled = True Then
            Set handler = New clsTextBox
            Set handler.TextBox = ctrl
            m_TextBoxHandlers.Add handler
        End If
    Next ctrl
    
    ' Re-enable and run initial calculation
    m_calculating = False
    Call HovercraftCalc.RunCalculations

End Sub

' ============================================================
' EVENT: UserForm_Activate
'
' Fires after the form becomes visible on screen.
' We use this rather than Initialize because the window
' handle needed for the API call is only available once
' the form is actually displayed.
' ============================================================
Private Sub UserForm_Activate()
    Call AddMinimizeButton
End Sub

' ============================================================
' SUB: AddMinimizeButton
'
' Adds a minimize button to the frmHovercraft title bar
' using Windows API calls to modify the form's window style.
'
' Called from UserForm_Activate in frmHovercraft.
'
' Note: We use UserForm_Activate rather than
' UserForm_Initialize because the window handle is not
' available until the form is fully visible on screen.
' ============================================================
Public Sub AddMinimizeButton()

    #If VBA7 Then
        Dim hwnd As LongPtr
        Dim lStyle As LongPtr
    #Else
        Dim lStyle As Long
    #End If

    ' Find the form's window handle using its caption text.
    ' "ThunderDFrame" is the internal Windows class name
    ' that Excel uses for all VBA UserForms.
    hwnd = FindWindowA("ThunderDFrame", frmHovercraft.Caption)

    If hwnd = 0 Then
        ' Window handle not found -- exit silently
        Exit Sub
    End If

    ' Read the current window style
    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE)

    ' Add the minimize box flag using bitwise OR.
    ' This adds WS_MINIMIZEBOX without changing any other styles.
    lStyle = lStyle Or WS_MINIMIZEBOX

    ' Write the updated style back to the window
    SetWindowLongPtr hwnd, GWL_STYLE, lStyle

    ' Redraw the title bar so the button appears immediately
    DrawMenuBar hwnd

End Sub

' ============================================================
' CHANGE EVENTS — Text Boxes
'
' Each text box on the form has a _Change event that fires
' every time the user types a character. We check the
' m_calculating flag first — if True, the change was made
' by our own code writing results back, and we ignore it.
' If False, the user made the change, so we recalculate.
'
' Rather than writing a separate handler for every text box,
' we route all of them to the same single line:
'   If Not m_calculating Then Call HovercraftCalc.RunCalculations
' ============================================================

Private Sub txtPa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtRho_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtWls_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtWdw_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtWtot_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtLCG_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtTCG_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtVCGf_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtVCGc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtLc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtBc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtAc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtDsk_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtDag_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtDc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtTheta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtPca_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtDeltaP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtCv_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtCc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtCd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtEta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub txtPf_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

' ============================================================
' CLICK EVENTS — Solve-For Radio Buttons
'
' When the user clicks a radio button to change the solve-for
' selection, we recalculate immediately so the field colors
' and values update right away.
' ============================================================

Private Sub optSolveWls_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveWdw_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveWtot_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveLc_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveBc_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveAc_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveDsk_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveDag_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveDc_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolvePca_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveCv_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveCc_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveCd_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveDeltaP_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolveEta_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optSolvePf_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub


' ============================================================
' CLICK EVENTS — Program-Wide Settings Radio Buttons
'
' Changing units or pressure mode triggers a recalculation
' so display values update immediately.
' ============================================================

Private Sub optUSCustomary_Click()
    ' Only act if we are actually switching from SI to US.
    ' If already in US mode, clicking this button does nothing.
    If Not m_calculating Then
        If Not isUS Then
            isUS = True
            Call HovercraftCalc.DisplayFromDictionary
            Call HovercraftCalc.UpdateFieldColors
        End If
    End If
End Sub

Private Sub optSI_Click()
    ' Only act if we are actually switching from US to SI.
    ' If already in SI mode, clicking this button does nothing.
    If Not m_calculating Then
        If isUS Then
            isUS = False
            Call HovercraftCalc.DisplayFromDictionary
            Call HovercraftCalc.UpdateFieldColors
        End If
    End If
End Sub

Private Sub optAbsolute_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub

Private Sub optGauge_Click()
    If Not m_calculating Then Call HovercraftCalc.RunCalculations
End Sub


' ============================================================
' CLICK EVENTS — Command Buttons
'
' These route to the corresponding Subs in HovercraftCalc.bas.
' Keeping the calculation logic OUT of the form module and IN
' the standard module is good practice — it separates the
' interface code from the business logic code.
' ============================================================

Private Sub cmdRecord_Click()
    ' Record current scenario to the Scenarios sheet
    Call HovercraftCalc.RecordScenario
End Sub

Private Sub cmdReset_Click()
    ' Reset all fields to default values
    Call HovercraftCalc.ResetForm
End Sub

Private Sub cmdClear_Click()
    ' Clear all input values
    Call HovercraftCalc.ClearForm
End Sub
