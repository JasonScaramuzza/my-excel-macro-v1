Attribute VB_Name = "modUtilities"

' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.1 (15 March 2026)
'
'  Module: [modUtilities]
'  Purpose: [Contains Utilities called by the other modules]
' ================================================================================


Option Explicit

' ============================================================
' TABLE HELPERS
' ============================================================

Public Function GetTableByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.Name) = LCase(tableName) Then
                Set GetTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function


Public Function GetCellByHeader(tbl As Object, rowRange As Range, headerName As String) As Variant
    Dim col As ListColumn
    Dim rowIndex As Long

    ' Guard: must be a ListObject
    If tbl Is Nothing Then
        GetCellByHeader = ""
        Exit Function
    End If

    If TypeName(tbl) <> "ListObject" Then
        GetCellByHeader = ""
        Exit Function
    End If

    ' Try to get the column
    On Error Resume Next
    Set col = tbl.ListColumns(headerName)
    On Error GoTo 0

    If col Is Nothing Then
        GetCellByHeader = ""
        Exit Function
    End If

    ' Calculate row index inside table
    rowIndex = rowRange.row - tbl.DataBodyRange.row + 1

    If rowIndex < 1 Or rowIndex > tbl.DataBodyRange.rows.Count Then
        GetCellByHeader = ""
        Exit Function
    End If

    GetCellByHeader = col.DataBodyRange.Cells(rowIndex, 1).Value
End Function


Public Function HeaderExistsInTable(tbl As ListObject, headerName As String) As Boolean
    Dim col As ListColumn
    
    If tbl Is Nothing Then Exit Function
    
    For Each col In tbl.ListColumns
        If StrComp(Trim$(col.Name), Trim$(headerName), vbTextCompare) = 0 Then
            HeaderExistsInTable = True
            Exit Function
        End If
    Next col
End Function


' ============================================================
' LOGGING (might want to add a RunID into this for each run)
' tblMovementLog headers:
'   MacroName | SourceTable | DestinationTable | TimeStamp |
'   ID | MovementType | Result | Details | Optional macroName
' ============================================================
Public Sub LogMovement(tblMovementLog As ListObject, _
                       ByVal sourceTableName As String, _
                       ByVal destTableName As String, _
                       ByVal pkValue As String, _
                       ByVal movementType As String, _
                       ByVal resultText As String, _
                       ByVal detailsText As String, _
                       Optional ByVal macroName As String = "RunWorkflow")
                       
    Dim newRow As ListRow
    Dim colMacro As Long, colSource As Long, colDest As Long
    Dim colTime As Long, colID As Long, colMoveType As Long
    Dim colResult As Long, colDetails As Long
    
    Set newRow = tblMovementLog.ListRows.Add
    
    ' Resolve columns by header (safer than hard-coded positions)
    colMacro = tblMovementLog.ListColumns("MacroName").Index
    colSource = tblMovementLog.ListColumns("SourceTable").Index
    colDest = tblMovementLog.ListColumns("DestinationTable").Index
    colTime = tblMovementLog.ListColumns("TimeStamp").Index
    colID = tblMovementLog.ListColumns("ID").Index
    colMoveType = tblMovementLog.ListColumns("MovementType").Index
    colResult = tblMovementLog.ListColumns("Result").Index
    colDetails = tblMovementLog.ListColumns("Details").Index
    

    newRow.Range.Cells(1, colMacro).Value = macroName
    newRow.Range.Cells(1, colSource).Value = sourceTableName
    newRow.Range.Cells(1, colDest).Value = destTableName
    newRow.Range.Cells(1, colTime).Value = Now
    newRow.Range.Cells(1, colID).Value = pkValue
    newRow.Range.Cells(1, colMoveType).Value = movementType
    newRow.Range.Cells(1, colResult).Value = resultText
    newRow.Range.Cells(1, colDetails).Value = detailsText
End Sub


' ============================================================
' UpdateLastMovedTable
'   Upserts LastPK and LastMovedAt for a SourceTable|DestinationTable
'   pair into an existing table named tblLastMoved.
'   Does NOT create the table. Returns True on success, False
'   if the table is missing or the update fails.
' ============================================================
Public Function UpdateLastMovedTable(ByVal sourceTableName As String, _
                                    ByVal destTableName As String, _
                                    ByVal pkValue As String) As Boolean
    Dim tbl As ListObject
    Dim keySource As String, keyDest As String
    Dim i As Long
    Dim lr As ListRow
    Dim ts As String

    UpdateLastMovedTable = False
    keySource = CStr(sourceTableName)
    keyDest = CStr(destTableName)
    ts = Format(Now, "yyyy-mm-dd HH:MM:SS")

    Set tbl = GetTableByName("tblLastMoved")
    If tbl Is Nothing Then Exit Function

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            If CStr(tbl.DataBodyRange.Cells(i, 1).Value) = keySource And _
               CStr(tbl.DataBodyRange.Cells(i, 2).Value) = keyDest Then
                tbl.DataBodyRange.Cells(i, 3).Value = pkValue
                tbl.DataBodyRange.Cells(i, 4).Value = ts
                UpdateLastMovedTable = True
                Exit Function
            End If
        Next i
    End If

    Set lr = tbl.ListRows.Add
    lr.Range.Cells(1, 1).Value = keySource
    lr.Range.Cells(1, 2).Value = keyDest
    lr.Range.Cells(1, 3).Value = pkValue
    lr.Range.Cells(1, 4).Value = ts

    UpdateLastMovedTable = True
End Function

' ============================================================
' PrimaryKeyExists
'
'   Checks whether a given primary key already exists in the
'   specified destination table.
'
'   RESPONSIBILITIES:
'   -----------------------------------------------
'   • Loop through all rows in the destination table.
'   • Read the Primary Key column for each row.
'   • Compare values using case-insensitive matching.
'   • Return True if a match is found, otherwise False.
'
'   PURPOSE:
'   -----------------------------------------------
'   • Supports duplication detection in PerformMovementForRow.
'   • Ensures "Create" movements do not generate duplicate
'     records in destination tables.
'   • Keeps duplication logic centralised and reusable.
'
'   NOTES:
'   -----------------------------------------------
'   • This function performs no logging.
'     Logging is handled by ExecuteMovementForID.
'
'   • This function assumes the destination table contains
'     a column named with a primary. Configuration validation
'     ensures this before workflow execution begins.
' ============================================================
Public Function PrimaryKeyExists(tbl As ListObject, pkValue As String, primaryKeyField As String) As Boolean
    Dim r As ListRow
    Dim idValue As Variant
    
    PrimaryKeyExists = False
    
    For Each r In tbl.ListRows
        idValue = GetCellByHeader(tbl, r.Range, primaryKeyField)
        
        If Trim(CStr(idValue)) <> "" Then
            If StrComp(Trim(CStr(idValue)), Trim(pkValue), vbTextCompare) = 0 Then
                PrimaryKeyExists = True
                Exit Function
            End If
        End If
    Next r
End Function

' ============================================================
' JoinCollection
'
' PURPOSE:
'   Utility function that concatenates all string items in a
'   VBA Collection into a single string, separated by the
'   specified delimiter.
'
' WHY THIS EXISTS:
'   VBA's built-in Join() only works on arrays, not Collections.
'   The workflow engine aggregates configuration-level fatal
'   errors into a Collection, so this helper produces the final
'   combined error message for logging.
'
' USAGE:
'   Used during configuration validation in RunWorkflow to
'   combine multiple fatal errors into one readable message
'   before logging a single "Aborted" entry.
'
' INPUTS:
'   col  - Collection containing string error messages.
'   sep  - Separator string (e.g. " | ").
'
' RETURNS:
'   A single concatenated string containing all items in col.
' ============================================================
Public Function JoinCollection(col As Collection, sep As String) As String
    Dim i As Long
    Dim result As String

    For i = 1 To col.Count
        If result = "" Then
            result = col(i)
        Else
            result = result & sep & col(i)
        End If
    Next i

    JoinCollection = result
End Function

' ============================================================
' GetPrimaryKeyForTable
'
' PURPOSE:
'   Returns the PrimaryKeyField for a given table name,
'   using tblWorkflowTables as the declarative source of truth.
'
' WHY THIS EXISTS:
'   Both the workflow engine and the import engine need to know
'   the primary key for each table. Hard-coding PK names causes
'   errors when table schemas evolve. This function centralises
'   PK lookup and keeps the architecture declarative.
'
' INPUTS:
'   tblWorkflowTables - ListObject containing:
'       Column 1: TableName
'       Column 2: PrimaryKeyField
'       Column 3: EligibleForWorkflow (Boolean)
'
'   tableName - Name of the table whose PK we want.
'
' RETURNS:
'   The primary key header name as a string, or "" if not found.
' ============================================================
Public Function GetPrimaryKeyForTable(tblWorkflowTables As ListObject, tableName As String) As String
    Dim r As ListRow

    If tblWorkflowTables Is Nothing Then Exit Function

    For Each r In tblWorkflowTables.ListRows
        If StrComp(Trim$(r.Range.Columns(1).Value), Trim$(tableName), vbTextCompare) = 0 Then
            GetPrimaryKeyForTable = Trim$(r.Range.Columns(2).Value)
            Exit Function
        End If
    Next r
End Function
