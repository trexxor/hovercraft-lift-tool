VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' CLASS MODULE: clsTextBox
'
' PURPOSE:
'   A reusable event handler that attaches to a single TextBox
'   and listens for keystrokes. One instance of this class is
'   created for each TextBox on the form during initialization.
'
' WHY A CLASS MODULE:
'   WithEvents can only be declared inside a class module.
'   This allows a single KeyDown handler to serve all TextBoxes
'   rather than writing individual handlers for each one.
' ============================================================

Option Explicit

' ----------------------------------------------------------
' WithEvents declaration
'
' "WithEvents" tells VBA that this variable is not just
' holding a reference to a TextBox -- it is actively
' listening to it. When the TextBox fires an event, VBA
' routes it through this variable to the handler below.
'
' This declaration makes the TextBox_KeyDown event available
' to write in this class, just like writing txtPa_KeyDown
' in the form module would be -- except this one handler
' works for ANY TextBox assigned to this variable.
' ----------------------------------------------------------
Public WithEvents TextBox As MSForms.TextBox
Attribute TextBox.VB_VarHelpID = -1


' ============================================================
' EVENT: TextBox_KeyDown
'
' Fires when the user presses any key while this TextBox
' has focus. We watch specifically for Enter (keycode 13)
' and Tab (keycode 9) and trigger a recalculation when
' either is pressed.
'
' Parameters:
'   KeyCode - the numeric code of the key that was pressed.
'             We can modify this value to suppress the
'             default key behavior if needed.
'   Shift   - indicates whether Shift, Ctrl, or Alt was
'             held. We do not use this here but it is
'             required by the event signature.
' ============================================================
Private Sub TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                            ByVal Shift As Integer)

    ' Check if the pressed key was Enter (13) or Tab (9)
    If KeyCode = 13 Or KeyCode = 9 Then

        ' Only recalculate if we are not already in the middle
        ' of a programmatic write-back -- same guard we use
        ' everywhere else in the form code
        If Not frmHovercraft.m_calculating Then
            Call HovercraftCalc.RunCalculations
        End If

    End If

End Sub

