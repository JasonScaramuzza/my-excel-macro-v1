Attribute VB_Name = "modMovementTypes"

' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.0 (14 March 2026)
'
'  Module: [modMovementTypes]
'  Purpose: [Contains the MovementType engine]
' ================================================================================

Option Explicit

' ============================================================
' ApplyMovementType (MovementType Engine)
'
' PURPOSE:
'   Centralises all MovementType behaviour (Create, Append, etc.)
'   in one place, mirroring how ApplyRule centralises RuleTypes etc.
'
'   - Takes the MovementType name and core movement context.
'   - Decides WHICH destination row to use (or whether movement
'     should not proceed).
'   - Returns:
'         • destRow (ListRow) when movement should continue.
'         • movementResult (ByRef) when movement should STOP
'           immediately (e.g. Duplicate, Not moved, or engine
'           not implemented).
'
' NOTES:
'   - This function does NOT apply WriteModes.
'   - It does NOT log.
'   - It does NOT handle transactional logic.
'   - Those responsibilities remain in PerformMovementForRow
'     and ExecuteMovementForID.
'
'   ENGINE PROBE:
'   - ValidateMovementTypeEngine will call this with dummy
'     arguments to ensure every MovementType marked Defined?=TRUE
'     has a corresponding Case block here.
' ============================================================
Public Function ApplyMovementType(ByVal movementType As String, _
                                  tblSource As ListObject, _
                                  tblDest As ListObject, _
                                  srcDataRow As ListRow, _
                                  ByVal pkValue As String, _
                                  ByVal primaryKeyField As String, _
                                  ByVal isPrimaryKeyField As Boolean, _
                                  ByRef movementResult As String) As ListRow
    
    Dim destRow As ListRow
    Dim isProbe As Boolean
    
    ' Default: no immediate result (caller will decide)
    movementResult = ""
    isProbe = (tblDest Is Nothing)   ' Probe mode when called from ValidateMovementTypeEngine

    Select Case movementType
    
        Case "Create"
            ' PROBE MODE:
            '   - We only care that this Case exists.
            '   - Do NOT touch any objects when probing.
            If isProbe Then
                Exit Function
            End If

            ' ------------------------------------------------
            ' CREATE :
            '
            '   If this mapping row is for the Primary Key field:
            '       - Check for duplicates in destination.
            '       - If duplicate found - stop movement and
            '         return "Duplicate".
            '
            '   For non-primary key fields:
            '       - Prefer existing row for this unique identifier (primary key).
            '       - If not found, create a new row.
            ' ------------------------------------------------
            If isPrimaryKeyField Then
                ' Only the Primary Key field is allowed to declare a true duplicate
                If PrimaryKeyExists(tblDest, pkValue, primaryKeyField) Then
                    movementResult = "Duplicate"
                    Set ApplyMovementType = Nothing
                    Exit Function
                End If
                Set destRow = GetOrCreateDestinationRow(tblDest, primaryKeyField)
            Else
                ' For non-primary key fields under Create:
                '   - Prefer existing row for this Unique Identifier (Primary Key)
                '   - If not found, create one (defensive)
                Set destRow = GetAppendDestinationRow(tblDest, pkValue, primaryKeyField)
                If destRow Is Nothing Then
                    Set destRow = GetOrCreateDestinationRow(tblDest, primaryKeyField)
                End If
            End If
            
            Set ApplyMovementType = destRow
        
        Case "Append"
            ' PROBE MODE:
            '   - Only need to know this Case exists.
            '   - Avoid touching any objects.
            If isProbe Then
                Exit Function
            End If

            ' ------------------------------------------------
            ' APPEND:
            '
            '   - Locate existing destination row by primary key.
            '   - If not found - movementResult = "Not moved"
            '     and caller will log accordingly.
            ' ------------------------------------------------
            Set destRow = GetAppendDestinationRow(tblDest, pkValue, primaryKeyField)
            If destRow Is Nothing Then
                movementResult = "Not moved"
                Set ApplyMovementType = Nothing
                Exit Function
            End If
            
            Set ApplyMovementType = destRow
        
        Case Else
            ' UNKNOWN MOVEMENTTYPE:
            '
            '   - In real runs, this should never happen because
            '     configuration validation + engine probe will
            '     catch it.
            '   - In probe mode, we self-report as not implemented.
            movementResult = "ENGINE_NOT_IMPLEMENTED"
            Set ApplyMovementType = Nothing
    End Select
End Function

' ============================================================
' PerformMovementForRow
'
'   Executes the actual movement logic for a single Primary Key
'   under a single mapping row.
'
'   RESPONSIBILITIES:
'   -----------------------------------------------
'   • Determine MovementType ("Create" or "Append").
'   • For "Create":
'         - Detect duplicates BEFORE creating a row.
'         - If duplicate - return "Duplicate" and skip.
'   • For "Append":
'         - Locate the existing destination row.
'         - If not found - return "Not moved".
'   • Select the correct destination row.
'   • Apply the WriteMode to the destination cell.
'   • Return a movement result string:
'         "Moved"
'         "Duplicate"
'         "Not moved"
'
'   NOTES:
'   -----------------------------------------------
'   • No logging is performed here.
'     Logging is handled centrally in ExecuteMovementForID.
'
'   • No validation is performed here.
'     All validation (RuleTypes, conditions, etc.) is
'     handled in ValidateRowForMappingRow.
'
'   • This function is intentionally low-level and
'     movement-focused. It should not make decisions
'     about whether a row *should* move — only *how*.
' ============================================================
Public Function PerformMovementForRow(tblSource As ListObject, _
                                      tblDest As ListObject, _
                                      srcDataRow As ListRow, _
                                      mappingRow As ListRow, _
                                      ByVal pkValue As String, _
                                      ByVal primaryKeyField As String) As String
    
    Dim movementType As String
    Dim sourceHeader As String
    Dim destHeader As String
    Dim writeMode As String
    Dim defaultValue As String
    
    Dim destRow As ListRow
    Dim destCell As Range
    Dim sourceValue As Variant
    Dim isPrimaryKeyField As Boolean
    Dim destColIndex As Long
    Dim movementResult As String   ' captures result from ApplyMovementType

    
    movementType = Trim(CStr(mappingRow.Range.Columns(3).Value))
    sourceHeader = Trim(CStr(mappingRow.Range.Columns(5).Value))
    destHeader = Trim(CStr(mappingRow.Range.Columns(6).Value))
    writeMode = Trim(CStr(mappingRow.Range.Columns(13).Value))
    defaultValue = Trim(CStr(mappingRow.Range.Columns(14).Value))
    
    ' primary key is table-driven
    isPrimaryKeyField = (LCase(destHeader) = LCase(primaryKeyField))
    
    ' --------------------------------------------------------
    ' Get source value
    '
    ' SourceValue:
    '   - If SourceHeader is set, read from source.
    '   - If SourceHeader is blank, treat sourceValue as "".
    '   - DefaultValue is only applied by specific WriteModes
    '     (e.g. FallbackDefaultValueIfBlank), never implicitly.
    ' --------------------------------------------------------
    If sourceHeader <> "" Then
        sourceValue = GetCellByHeader(tblSource, srcDataRow.Range, sourceHeader)
    Else
        sourceValue = ""   ' no implicit defaulting when no SourceHeader
    End If

    

    
    ' ========================================================
    ' MODIFIED: MovementType now resolved via ApplyMovementType
    '
    '   - Removed hard-coded Select Case on movementType.
    '   - All MovementType-specific behaviour (Create/Append)
    '     is now centralised in ApplyMovementType.
    '   - ApplyMovementType returns:
    '         • destRow when movement should continue.
    '         • movementResult ("Duplicate", "Not moved",
    '           "ENGINE_NOT_IMPLEMENTED") when movement
    '           should stop immediately.
    ' ========================================================
    Set destRow = ApplyMovementType(movementType, tblSource, tblDest, srcDataRow, pkValue, primaryKeyField, isPrimaryKeyField, movementResult)
    
    ' If the engine signalled an immediate result, return it
    If movementResult <> "" Then
        PerformMovementForRow = movementResult
        Exit Function
    End If
    
    ' Safety: if no destination row was returned, treat as Not moved
    If destRow Is Nothing Then
        PerformMovementForRow = "Not moved"
        Exit Function
    End If
    ' ========================================================
    
    ' --------------------------------------------------------
    ' If no destination header (e.g. pure timestamp row with blank SourceHeader), skip cell write
    ' --------------------------------------------------------
    If destHeader = "" Then
        PerformMovementForRow = "Moved"
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' Get destination cell
    ' --------------------------------------------------------
    destColIndex = tblDest.ListColumns(destHeader).Index
    Set destCell = destRow.Range.Cells(1, destColIndex)
    
    ' --------------------------------------------------------
    ' Apply WriteMode
    ' --------------------------------------------------------
    '      Pass DefaultValue into the WriteMode engine so that
    '      modes like FallbackDefaultValueIfBlank can use it
    '      explicitly and declaratively.
    ApplyWriteMode writeMode, sourceValue, defaultValue, destCell
    
    PerformMovementForRow = "Moved"
End Function


' ============================================================
' GetOrCreateDestinationRow (for Create)
'   Reuses an existing blank Primary Key row if available,
'   otherwise adds a new row.
' ============================================================
Public Function GetOrCreateDestinationRow(tblDest As ListObject, _
                                          ByVal primaryKeyField As String) As ListRow
    
    Dim r As ListRow
    Dim IDColIndex As Long
    Dim IDVal As String
    
    IDColIndex = tblDest.ListColumns(primaryKeyField).Index
    
    ' Look for existing blank Primary Key row
    For Each r In tblDest.ListRows
        IDVal = Trim(CStr(r.Range.Cells(1, IDColIndex).Value))
        If IDVal = "" Then
            Set GetOrCreateDestinationRow = r
            Exit Function
        End If
    Next r
    
    ' None found - add new row
    Set GetOrCreateDestinationRow = tblDest.ListRows.Add
End Function


' ============================================================
' GetAppendDestinationRow (for Append)
'   Finds existing row in destination by primary key
' ============================================================
Public Function GetAppendDestinationRow(tblDest As ListObject, _
                                        ByVal pkValue As String, _
                                        ByVal primaryKeyField As String) As ListRow
    
    Dim rng As Range
    Dim pkCol As ListColumn
    
    On Error Resume Next
    Set pkCol = tblDest.ListColumns(primaryKeyField)
    On Error GoTo 0
    
    If pkCol Is Nothing Then Exit Function
    
    On Error Resume Next
    Set rng = pkCol.DataBodyRange.Find(What:=pkValue, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If rng Is Nothing Then
        Set GetAppendDestinationRow = Nothing
    Else
        Set GetAppendDestinationRow = tblDest.ListRows(rng.row - tblDest.DataBodyRange.row + 1)
    End If
End Function

