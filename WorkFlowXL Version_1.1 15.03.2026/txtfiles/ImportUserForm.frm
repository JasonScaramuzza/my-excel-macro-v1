VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportUserForm 
   Caption         =   "ImportUserForm"
   ClientHeight    =   3030
   ClientLeft      =   -40
   ClientTop       =   -300
   ClientWidth     =   5840
   OleObjectBlob   =   "ImportUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.1 (15 March 2026)
'
'  Form: [ImportUserForm]
'  Purpose: [Contains the ImportUserform that calls RunImport]
' ================================================================================

Option Explicit

' ============================================================
' USERFORM INITIALISATION
' ============================================================
Private Sub UserForm_Initialize()

    Dim wb As Workbook
    Dim ctrl As Control

    ' --------------------------------------------------------
    ' Prevent shrinking / enforce consistent size
    ' --------------------------------------------------------
    Me.Width = 420
    Me.Height = 260

    ' --------------------------------------------------------
    ' Normalise font sizes for all controls
    ' --------------------------------------------------------
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Font.Size = 10
        On Error GoTo 0
    Next ctrl

    ' --------------------------------------------------------
    ' Explicit control layout (prevents DPI shrink issues)
    ' --------------------------------------------------------
    With Me.lblWorkbook
        .Caption = "Select Workbook:"
        .Left = 20
        .Top = 20
        .Width = 120
        .Height = 18
    End With

    With Me.cboWorkbook
        .Left = 20
        .Top = 40
        .Width = 260
        .Height = 22
    End With

    With Me.lblWorksheet
        .Caption = "Select Worksheet:"
        .Left = 20
        .Top = 75
        .Width = 120
        .Height = 18
    End With

    With Me.cboWorksheet
        .Left = 20
        .Top = 95
        .Width = 260
        .Height = 22
    End With

    With Me.lblUniqueID
        .Caption = "Optional UniqueID Override:"
        .Left = 20
        .Top = 130
        .Width = 200
        .Height = 18
    End With

    With Me.txtUniqueID
        .Left = 20
        .Top = 150
        .Width = 260
        .Height = 22
    End With

    With Me.btnImport
        .Caption = "Import"
        .Left = 300
        .Top = 40
        .Width = 90
        .Height = 30
    End With

    With Me.btnCancel
        .Caption = "Cancel"
        .Left = 300
        .Top = 85
        .Width = 90
        .Height = 30
    End With


    ' --------------------------------------------------------
    ' Populate workbook dropdown
    ' --------------------------------------------------------
    Me.cboWorkbook.Clear

    For Each wb In Application.Workbooks
        Me.cboWorkbook.AddItem wb.Name
    Next wb

End Sub


' ============================================================
' WORKBOOK SELECTION ? POPULATE SHEETS
' ============================================================
Private Sub cboWorkbook_Change()

    Dim ws As Worksheet
    Dim wb As Workbook

    Me.cboWorksheet.Clear

    If Me.cboWorkbook.Value = "" Then Exit Sub

    On Error GoTo ErrHandler
    Set wb = Application.Workbooks(Me.cboWorkbook.Value)

    For Each ws In wb.Sheets
        Me.cboWorksheet.AddItem ws.Name
    Next ws

    Exit Sub

ErrHandler:
    MsgBox "Error: Unable to access workbook '" & Me.cboWorkbook.Value & "'.", vbCritical

End Sub


' ============================================================
' CANCEL BUTTON
' ============================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub


' ============================================================
' MAIN IMPORT BUTTON
' ============================================================
Private Sub btnImport_Click()

    Dim selectedWB As Workbook
    Dim selectedWS As Worksheet
    Dim userProvidedID As String

    ' --------------------------------------------------------
    ' Validate selections
    ' --------------------------------------------------------
    If Me.cboWorkbook.Value = "" Or Me.cboWorksheet.Value = "" Then
        MsgBox "Please select both a workbook and a sheet before importing.", vbExclamation
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' Resolve workbook + sheet
    ' --------------------------------------------------------
    Set selectedWB = Application.Workbooks(Me.cboWorkbook.Value)
    Set selectedWS = selectedWB.Sheets(Me.cboWorksheet.Value)

    ' --------------------------------------------------------
    ' Optional UniqueID override
    ' --------------------------------------------------------
    userProvidedID = Trim(Me.txtUniqueID.Value)

    ' --------------------------------------------------------
    ' Call the new import engine
    ' --------------------------------------------------------
    Call StartImportFromUserform(selectedWB, selectedWS, userProvidedID)
    
    ' --------------------------------------------------------
    ' Only unload AFTER the import completes successfully
    ' --------------------------------------------------------
    Unload Me


End Sub
