VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================
' THISWORKBOOK MODULE
'
' PURPOSE:
'   Code attached to the workbook itself (not a sheet, not a
'   standard module). The Workbook_Open event fires
'   automatically every time the file is opened.
'
' HOW TO INSTALL THIS CODE:
'   1. Press Alt+F11 to open the VBA Editor
'   2. In the Project panel, double-click "ThisWorkbook"
'      (under Microsoft Excel Objects)
'   3. Paste this code into that window
' ============================================================

Option Explicit

' ============================================================
' EVENT: Workbook_Open
'
' Excel calls this automatically when the workbook is opened.
' We use it to:
'   1. Navigate to the Launch sheet so the user sees the button
'   2. Show the calculator form in modeless mode
'
' "Modeless" means the form stays open and visible while the
' user can still interact with the spreadsheet underneath.
' This is achieved by passing the vbModeless argument to .Show
' instead of the default vbModal (which would block the sheet).
' ============================================================
Private Sub Workbook_Open()

    ' Navigate to the Launch sheet
    ThisWorkbook.Sheets("Launch").Activate

    ' Load and show the form in modeless mode
    ' vbModeless = 0 (the form stays open alongside the sheet)
    frmHovercraft.Show vbModeless

End Sub

' ============================================================
' EVENT: btnLaunch_Click
'
' Called when the user clicks the "Launch Hovercraft Calculator"
' button on the Launch sheet.
'
' Shows the calculator form in modeless mode, meaning the form
' stays open and visible while the user can still interact with
' the spreadsheet underneath it.
'
' Note: If the form is already open, calling .Show again will
' simply bring it to the front rather than opening a second
' instance.
' ============================================================
Sub btnLaunch_Click()

    frmHovercraft.Show vbModeless

End Sub
