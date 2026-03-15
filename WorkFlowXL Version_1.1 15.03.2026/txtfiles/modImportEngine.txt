Attribute VB_Name = "modImportEngine"

' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.1 (15 March 2026)
'
'  Module: [modImportEngine]
'  Purpose: [Contains the ImportEngine]
' ================================================================================

' ================================================================================
' Please Note this is not necessarily as useful as the other modules at this stage,
' due to the nature of this copying from an external sheet. In future versions I may
' update this so it looks for the headers on the external sheet in a better way instead
' of assuming they start in A1
' ================================================================================

Option Explicit

' ================================================================================
'  WorkflowXL – Import Engine
'
'  PURPOSE:
'      Imports data from an external workbook into tblStaging using tblMap_Import.
'      Fully table-driven, no hardcoding, aligned with WorkflowXL architecture.
'
'  FEATURES:
'      • Validates RuleType engine before import
'      • Validates each row using RuleTypes from tblMap_Import
'      • Duplicate detection across ALL workflow tables (EligibleForWorkflow = TRUE)
'      • Logs all results using WorkflowXL LogMovement format
'      • Updates tblLastMoved using UpdateLastMovedTable
'      • Uses generic GetLastMovedPK helper
'      • Uses starting-point logic:
'            1) User override
'            2) tblLastMoved
'            3) Last N rows (default 100)
'
'  ENTRY POINTS:
'      • RunImport
'      • StartImportFromUserform (called by userform)
'
' ================================================================================


' ============================================================
' USERFORM ENTRY POINT
' ============================================================
Public Sub StartImportFromUserform(selectedWB As Workbook, selectedWS As Worksheet, userProvidedID As String)
    Call RunImport(selectedWB, selectedWS, userProvidedID)
End Sub



' ============================================================
' MAIN IMPORT MACRO
' ============================================================
Public Sub RunImport(selectedWB As Workbook, selectedWS As Worksheet, Optional userProvidedID As String = "")

    Const DEFAULT_SEARCH_WINDOW As Long = 100
    Const SOURCE_TABLE_NAME As String = "External Workbook"
    Const DEST_TABLE_NAME As String = "tblStaging"

    Dim wsConfig As Worksheet
    Dim wsLogs As Worksheet
    Dim wsStaging As Worksheet

    Dim tblStaging As ListObject
    Dim tblMap As ListObject
    Dim tblRuleList As ListObject
    Dim tblMovementLog As ListObject
    Dim tblLastMoved As ListObject
    Dim tblWorkflowTables As ListObject
    
    Dim stagingPK As String

    Dim mapDict As Object
    Dim ruleDict As Object

    Dim configErrors As New Collection
    Dim ruleEngineError As String

    Dim lastImportedID As String
    Dim startRow As Long
    Dim srcRow As Long
    Dim lastRow As Long

    Dim logicalField As Variant
    Dim entry As Object
    Dim uniqueID As String
    Dim lastSuccessfulID As String

    Dim failures As Collection
    Dim detailsText As String
    Dim resultText As String

    Dim newRow As ListRow


    ' --------------------------------------------------------
    ' Locate required sheets and tables
    ' --------------------------------------------------------
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("Configuration")
    Set wsLogs = ThisWorkbook.Worksheets("Movement Logs")
    Set wsStaging = ThisWorkbook.Worksheets("Staging Sheet")

    Set tblStaging = wsStaging.ListObjects("tblStaging")
    Set tblMap = wsConfig.ListObjects("tblMap_Import")
    Set tblRuleList = wsConfig.ListObjects("tblRuleList")
    Set tblMovementLog = wsLogs.ListObjects("tblMovementLog")
    Set tblLastMoved = wsLogs.ListObjects("tblLastMoved")
    Set tblWorkflowTables = wsConfig.ListObjects("tblWorkflowTables")
    On Error GoTo 0

    ' --------------------------------------------------------
    ' Validate required tables exist
    ' --------------------------------------------------------
    If tblStaging Is Nothing Then configErrors.Add "tblStaging not found."
    If tblMap Is Nothing Then configErrors.Add "tblMap_Import not found."
    If tblRuleList Is Nothing Then configErrors.Add "tblRuleList not found."
    If tblMovementLog Is Nothing Then configErrors.Add "tblMovementLog not found."
    If tblLastMoved Is Nothing Then configErrors.Add "tblLastMoved not found."
    If tblWorkflowTables Is Nothing Then configErrors.Add "tblWorkflowTables not found."

    If configErrors.Count > 0 Then
        Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, "", "Import", "Aborted", JoinCollection(configErrors, " | "), "RunImport")
        MsgBox "Import aborted due to configuration errors:" & vbCrLf & JoinCollection(configErrors, vbCrLf), vbCritical
        Exit Sub
    End If

    stagingPK = GetPrimaryKeyForTable(tblWorkflowTables, "tblStaging")

    ' --------------------------------------------------------
    ' Build RuleType dictionary
    ' --------------------------------------------------------
    Set ruleDict = BuildRuleDictionary(tblRuleList)

    ' --------------------------------------------------------
    ' Validate RuleType engine
    ' --------------------------------------------------------
    ruleEngineError = ValidateRuleEngine(ruleDict)
    If ruleEngineError <> "" Then
        Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, "", "Import", "Aborted", ruleEngineError, "RunImport")
        MsgBox "Import aborted: " & ruleEngineError, vbCritical
        Exit Sub
    End If


    ' --------------------------------------------------------
    ' Load mapping dictionary
    ' --------------------------------------------------------
    Set mapDict = LoadImportMappingFromTable(tblMap)

    ' Ensure UniqueID exists in mapping
    If Not mapDict.Exists("UniqueID") Then
        Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, "", "Import", "Aborted", "LogicalField 'UniqueID' missing in tblMap_Import.", "RunImport")
        MsgBox "Import aborted: LogicalField 'UniqueID' missing in tblMap_Import.", vbCritical
        Exit Sub
    End If


    ' --------------------------------------------------------
    ' Determine starting UniqueID
    ' --------------------------------------------------------
    lastImportedID = GetLastMovedPK(SOURCE_TABLE_NAME, DEST_TABLE_NAME)

    startRow = FindStartingRowForImport( _
                    selectedWS, _
                    mapDict("UniqueID"), _
                    userProvidedID, _
                    lastImportedID, _
                    DEFAULT_SEARCH_WINDOW)

    If startRow = 0 Then
        Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, "", "Import", "Aborted", _
                         "Unable to determine starting row for import.", "RunImport")
        MsgBox "Import aborted: Unable to determine starting row.", vbCritical
        Exit Sub
    End If


    ' --------------------------------------------------------
    ' MAIN IMPORT LOOP
    ' --------------------------------------------------------
    lastRow = selectedWS.Cells(selectedWS.rows.Count, 1).End(xlUp).row

    For srcRow = startRow To lastRow

        Set failures = New Collection
        detailsText = ""
        resultText = ""

        ' Extract UniqueID early
        uniqueID = ExtractSourceValue(selectedWS, srcRow, mapDict("UniqueID"))

        If Trim(uniqueID) = "" Then
            Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, "", "Import", "Not imported", "Missing UniqueID", "RunImport")
            GoTo NextRow
        End If

        ' Duplicate detection across workflow
        If UniqueIDExistsInWorkflow(uniqueID, tblStaging, tblWorkflowTables) Then
            Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, uniqueID, "Import", "Not imported", "Duplicate UniqueID", "RunImport")
            GoTo NextRow
        End If

        ' Validate row using RuleTypes
        Call ValidateImportRow(selectedWS, srcRow, mapDict, ruleDict, failures)

        If failures.Count > 0 Then
            detailsText = JoinCollection(failures, " | ")
            Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, uniqueID, "Import", "Not imported", detailsText, "RunImport")
            GoTo NextRow
        End If

        ' Write row to tblStaging
        Set newRow = WriteImportRowToTable(selectedWS, srcRow, mapDict, tblStaging)

        lastSuccessfulID = uniqueID

        Call LogMovement(tblMovementLog, SOURCE_TABLE_NAME, DEST_TABLE_NAME, uniqueID, "Import", "Imported", "", "RunImport")

NextRow:
    Next srcRow


    ' --------------------------------------------------------
    ' Update LastMoved table
    ' --------------------------------------------------------
    If lastSuccessfulID <> "" Then
        Call UpdateLastMovedTable(SOURCE_TABLE_NAME, DEST_TABLE_NAME, lastSuccessfulID)
    End If

    MsgBox "Import complete.", vbInformation

End Sub



' ============================================================
' GENERIC: Get last moved PK from tblLastMoved
' ============================================================
Public Function GetLastMovedPK(ByVal sourceTableName As String, ByVal destTableName As String) As String
    Dim tbl As ListObject
    Dim r As ListRow

    Set tbl = GetTableByName("tblLastMoved")
    If tbl Is Nothing Then Exit Function

    For Each r In tbl.ListRows
        If Trim(CStr(r.Range.Cells(1, 1).Value)) = sourceTableName And _
           Trim(CStr(r.Range.Cells(1, 2).Value)) = destTableName Then

            GetLastMovedPK = Trim(CStr(r.Range.Cells(1, 3).Value))
            Exit Function
        End If
    Next r
End Function



' ============================================================
' FIND STARTING ROW FOR IMPORT
' ============================================================
Private Function FindStartingRowForImport(ws As Worksheet, _
                                          entryUniqueID As Object, _
                                          userProvidedID As String, _
                                          lastImportedID As String, _
                                          searchWindow As Long) As Long

    Dim srcCol As Long
    Dim r As Long
    Dim lastRow As Long
    Dim val As String

    srcCol = FindSourceColumn(ws, entryUniqueID)
    If srcCol = 0 Then Exit Function

    lastRow = ws.Cells(ws.rows.Count, srcCol).End(xlUp).row

    ' 1) User override
    If Trim(userProvidedID) <> "" Then
        For r = lastRow To 2 Step -1
            If Trim(CStr(ws.Cells(r, srcCol).Value)) = Trim(userProvidedID) Then
                FindStartingRowForImport = r
                Exit Function
            End If
        Next r
    End If

    ' 2) LastMovedPK
    If Trim(lastImportedID) <> "" Then
        For r = lastRow To 2 Step -1
            If Trim(CStr(ws.Cells(r, srcCol).Value)) = Trim(lastImportedID) Then
                FindStartingRowForImport = r + 1
                Exit Function
            End If
        Next r
    End If

    ' 3) Search last N rows for first non-blank UniqueID ***This doesn't really work as intended at the moment
    For r = lastRow To Application.Max(2, lastRow - searchWindow) Step -1
        val = Trim(CStr(ws.Cells(r, srcCol).Value))
        If val <> "" Then
            FindStartingRowForImport = r
            Exit Function
        End If
    Next r
End Function



' ============================================================
' DUPLICATE DETECTION ACROSS WORKFLOW (Declarative PK Version)
'
' PURPOSE:
'   Checks whether a UniqueID already exists in:
'       • tblStaging
'       • Any workflow table marked EligibleForWorkflow = TRUE
'
'   Uses tblWorkflowTables to determine the correct PK field
'   for each table, including tblStaging.
'
' WHY THIS EXISTS:
'   Prevents duplicate imports and ensures consistency with
'   the workflow engine's declarative PK architecture.
' ============================================================
Private Function UniqueIDExistsInWorkflow(uniqueID As String, _
                                          tblStaging As ListObject, _
                                          tblWorkflowTables As ListObject) As Boolean

    Dim r As ListRow
    Dim tbl As ListObject
    Dim tableName As String
    Dim pkField As String

    ' --------------------------------------------------------
    ' 1) Check tblStaging using its declarative PK
    ' --------------------------------------------------------
    pkField = GetPrimaryKeyForTable(tblWorkflowTables, "tblStaging")

    If pkField <> "" Then
        If PrimaryKeyExists(tblStaging, uniqueID, pkField) Then
            UniqueIDExistsInWorkflow = True
            Exit Function
        End If
    End If

    ' --------------------------------------------------------
    ' 2) Check all workflow tables marked EligibleForWorkflow
    ' --------------------------------------------------------
    For Each r In tblWorkflowTables.ListRows
        If CBool(r.Range.Columns(3).Value) = True Then   ' EligibleForWorkflow
            tableName = Trim$(r.Range.Columns(1).Value)
            pkField = Trim$(r.Range.Columns(2).Value)

            Set tbl = GetTableByName(tableName)
            If Not tbl Is Nothing Then
                If PrimaryKeyExists(tbl, uniqueID, pkField) Then
                    UniqueIDExistsInWorkflow = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function



' ============================================================
' VALIDATE IMPORT ROW USING RULETYPES
' ============================================================
Private Sub ValidateImportRow(ws As Worksheet, srcRow As Long, mapDict As Object, ruleDict As Object, failures As Collection)

    Dim logicalField As Variant
    Dim entry As Object
    Dim srcCol As Long
    Dim val As Variant
    Dim errMsg As String

    For Each logicalField In mapDict.Keys

        Set entry = mapDict(logicalField)
        If entry("Active") = False Then GoTo NextField

        srcCol = FindSourceColumn(ws, entry)
        If srcCol = 0 Then
            failures.Add logicalField & ": Source header not found"
            GoTo NextField
        End If

        val = ws.Cells(srcRow, srcCol).Value

        errMsg = ApplyRule(entry("RuleType"), CStr(logicalField), val)
        If errMsg <> "" Then failures.Add errMsg

NextField:
    Next logicalField
End Sub



' ============================================================
' WRITE IMPORT ROW TO tblStaging
' ============================================================
Private Function WriteImportRowToTable(ws As Worksheet, srcRow As Long, mapDict As Object, tblStaging As ListObject) As ListRow

    Dim newRow As ListRow
    Dim logicalField As Variant
    Dim entry As Object
    Dim srcCol As Long
    Dim destCol As Long

    Set newRow = tblStaging.ListRows.Add

    For Each logicalField In mapDict.Keys

        Set entry = mapDict(logicalField)
        If entry("Active") = False Then GoTo NextField

        srcCol = FindSourceColumn(ws, entry)
        If srcCol > 0 Then
            destCol = tblStaging.ListColumns(entry("DestinationHeader")).Index
            newRow.Range(1, destCol).Value = ws.Cells(srcRow, srcCol).Value
        End If

NextField:
    Next logicalField

    Set WriteImportRowToTable = newRow
End Function



' ============================================================
' SAFE SOURCE VALUE EXTRACTION
' ============================================================
Private Function ExtractSourceValue(ws As Worksheet, rowNum As Long, entry As Object) As Variant
    Dim col As Long
    col = FindSourceColumn(ws, entry)
    If col = 0 Then
        ExtractSourceValue = ""
    Else
        ExtractSourceValue = ws.Cells(rowNum, col).Value
    End If
End Function

'------------------------------------------------------------
' LoadImportMappingFromTable
'   Reads tblMap_Import and returns a dictionary:
'   LogicalField ? mapping object:
'       .SourceA
'       .SourceB
'       .Aliases (dictionary)
'       .DestinationHeader
'       .Active
'       .RuleType
'------------------------------------------------------------
Public Function LoadImportMappingFromTable(tblMap As ListObject) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As ListRow
    Dim entry As Object
    Dim aliasDict As Object
    Dim aliasList As String
    Dim aliasArr As Variant
    Dim a As Variant
    Dim key As String

    For Each r In tblMap.ListRows

        ' Create the entry dictionary for this row
        Set entry = CreateObject("Scripting.Dictionary")

        entry("LogicalField") = Trim(CStr(r.Range.Columns(1).Value))
        entry("SourceA") = Trim(CStr(r.Range.Columns(2).Value))
        entry("SourceB") = Trim(CStr(r.Range.Columns(3).Value))
        aliasList = Trim(CStr(r.Range.Columns(4).Value))
        entry("DestinationHeader") = Trim(CStr(r.Range.Columns(5).Value))
        entry("Active") = CBool(r.Range.Columns(6).Value)
        entry("RuleType") = Trim(CStr(r.Range.Columns(7).Value))

        ' Build alias dictionary
        Set aliasDict = CreateObject("Scripting.Dictionary")
        If aliasList <> "" Then
            aliasArr = Split(aliasList, ",")
            For Each a In aliasArr
                aliasDict(Trim(LCase(CStr(a)))) = True
            Next a
        End If
        entry.Add "Aliases", aliasDict

        ' ------------------------------
        ' VALIDATION BEFORE ADDING ENTRY
        ' ------------------------------
        key = entry("LogicalField")

        ' Skip inactive rows
        If entry("Active") = False Then GoTo NextRow

        ' Skip rows missing required fields
        If key = "" Then GoTo NextRow
        If entry("DestinationHeader") = "" Then GoTo NextRow
        If entry("RuleType") = "" Then GoTo NextRow

        ' Prevent duplicate keys
        If dict.Exists(key) Then
            MsgBox "Duplicate LogicalField detected in tblMap_Import: " & key, vbCritical
            GoTo NextRow
        End If

        ' Safe to add
        dict.Add key, entry

NextRow:
    Next r

    Set LoadImportMappingFromTable = dict
End Function


'------------------------------------------------------------
' FindSourceColumn
'   Matches a source header using:
'       Source A Header
'       Source B Header
'       Aliases
'   Returns column index or 0 if not found.
'------------------------------------------------------------
Function FindSourceColumn(ws As Worksheet, entry As Object) As Long
    Dim lastCol As Long, c As Long
    Dim hdr As String
    Dim cleaned As String
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        hdr = Trim(CStr(ws.Cells(1, c).Value))
        cleaned = LCase(hdr)
        
        ' Match Source A
        If cleaned = LCase(entry("SourceA")) Then
            FindSourceColumn = c
            Exit Function
        End If
        
        ' Match Source B
        If cleaned = LCase(entry("SourceB")) Then
            FindSourceColumn = c
            Exit Function
        End If
        
        ' Match Aliases
        If entry("Aliases").Exists(cleaned) Then
            FindSourceColumn = c
            Exit Function
        End If
    Next c
    
    FindSourceColumn = 0
End Function


