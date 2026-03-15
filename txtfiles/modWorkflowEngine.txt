Attribute VB_Name = "modWorkflowEngine"


' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.0 (14 March 2026)
'
'  Module: [modWorkFlowEngine]
'  Purpose: [Contains the main WorkFlowEngine]
' ================================================================================


Option Explicit

Private duplicateLogDict As Object

' ============================================================
' RunWorkflow
'
'   Top-level orchestrator for the entire workflow engine.
'
'   RESPONSIBILITIES:
'   -----------------------------------------------
'   • Locate all required configuration tables:
'         - tblWorkflowMapping
'         - tblRuleList
'         - tblMovementTypes
'         - tblMovementLog
'
'   • Validate each mapping row BEFORE processing any data:
'         - Missing tables
'         - Missing headers
'         - Invalid RuleTypes
'         - Invalid MovementTypes
'         - Invalid WriteModes
'
'   • If a configuration error is found:
'         - Log ONE "Aborted" entry (no Unique Identifier (primary Key))
'         - Display a message
'         - Stop the workflow immediately
'
'   • Build dictionaries:
'         - RuleType dictionary
'         - MovementType dictionary
'
'   • Collect mapping rows and sort by ExecutionOrder.
'
'   • For each mapping row:
'         - Call ExecuteMovementForID to process all Unique Identifiers (Primary Key values).
'
'   LOGGING:
'   -----------------------------------------------
'   • Configuration errors - ONE log entry (no Primary Key Value).
'   • Data-row movement - handled inside ExecuteMovementForID.
'
'   PURPOSE:
'   -----------------------------------------------
'   Provides a clean, predictable, top-down execution flow
'   for the entire workflow engine.
' ============================================================

Public Sub RunWorkflow(Optional ByVal SourceTableFilter As String = "")
    Dim wsConfig As Worksheet
    Dim wsLogs As Worksheet
    Dim tblMap As ListObject
    Dim tblMovementLog As ListObject
    Dim tblLastMoved As ListObject
    
    Dim fatalErrors As New Collection
    
    Dim tblRuleList As ListObject
    Dim ruleDict As Object
    
    Dim tblMovementTypes As ListObject
    Dim movementTypeDict As Object
    
    Dim tblConditionalOperators As ListObject
    Dim operatorDict As Object
    
    Dim tblWriteModes As ListObject
    Dim writeModeDict As Object
    
    Dim tblWorkflowTables As ListObject
    Dim tablePKDict As Object
    Dim tableEligibleDict As Object
    Dim tableExecOrderDict As Object
    
    Dim mappingRows As Collection
    Dim mappingRow As ListRow
    Dim configError As String
    
    Dim sourceTableName As String
    Dim destTableName As String
    Dim movementType As String
    Dim execOrder As Long
    
    Dim groupedMappings As Object   ' Dictionary: key = "Source|Dest", value = Collection of mapping rows
    Dim groupKey As Variant
    Dim groupRows As Collection
    
        
    On Error GoTo ErrHandler
    
    ' --------------------------------------------------------
    ' Locate core sheets and tables
    ' --------------------------------------------------------
    Set wsConfig = ThisWorkbook.Worksheets("Configuration")
    Set wsLogs = ThisWorkbook.Worksheets("Movement Logs")
    
    On Error Resume Next
    Set tblMap = wsConfig.ListObjects("tblWorkflowMapping")
    Set tblMovementLog = wsLogs.ListObjects("tblMovementLog")
    Set tblLastMoved = wsLogs.ListObjects("tblLastMoved")
    Set tblRuleList = wsConfig.ListObjects("tblRuleList")
    Set tblMovementTypes = wsConfig.ListObjects("tblMovementTypes")
    Set tblConditionalOperators = wsConfig.ListObjects("tblConditionalOperators")
    Set tblWriteModes = wsConfig.ListObjects("tblWriteModes")
    Set tblWorkflowTables = wsConfig.ListObjects("tblWorkflowTables")
    On Error GoTo ErrHandler
    
    ' Initialise duplicate suppression dictionary for THIS run
    Set duplicateLogDict = CreateObject("Scripting.Dictionary")
    
    If tblMap Is Nothing Then
        MsgBox "Workflow aborted: tblWorkflowMapping not found.", vbCritical
        Exit Sub
    End If
    
    If tblMovementLog Is Nothing Then
        MsgBox "Workflow aborted: tblMovementLog not found.", vbCritical
        Exit Sub
    End If
    
    If tblRuleList Is Nothing Then
        MsgBox "Workflow aborted: tblRuleList not found.", vbCritical
        Exit Sub
    End If
    
    If tblMovementTypes Is Nothing Then
        MsgBox "Workflow aborted: tblMovementTypes not found.", vbCritical
        Exit Sub
    End If
    
    If tblConditionalOperators Is Nothing Then
        MsgBox "Workflow aborted: tblConditionalOperators not found.", vbCritical
        Exit Sub
    End If
    
    If tblWriteModes Is Nothing Then
        MsgBox "Workflow aborted: tblWriteModes not found.", vbCritical
        Exit Sub
    End If

    If tblWorkflowTables Is Nothing Then
        MsgBox "Workflow aborted: tblWorkflowTables not found.", vbCritical
        Exit Sub
    End If
    
    If tblLastMoved Is Nothing Then
        MsgBox "Workflow aborted: tblLastMoved not found.", vbCritical
        Exit Sub
    End If
    
    ' --------------------------------------------------------
    ' Build dictionaries
    ' --------------------------------------------------------
    Set ruleDict = BuildRuleDictionary(tblRuleList)
    
    Set operatorDict = BuildConditionalOperatorDictionary(tblConditionalOperators)
    
    Set movementTypeDict = BuildMovementTypeDictionary(tblMovementTypes)
    
    Set writeModeDict = BuildWriteModeDictionary(tblWriteModes)
    
    ' --------------------------------------------------------
    ' Build table-level configuration dictionaries
    ' --------------------------------------------------------
    Call BuildTableConfigurationDictionaries(tblWorkflowTables, _
                                         tablePKDict, _
                                         tableEligibleDict, _
                                         tableExecOrderDict)
    
    
    ' Validate that every RuleType in the rule table is implemented in the rule engine
    Dim ruleEngineError As String
    ruleEngineError = ValidateRuleEngine(ruleDict)
    
    If ruleEngineError <> "" Then
        fatalErrors.Add ruleEngineError
    End If
    
    
    ' --------------------------------------------------------
    'MovementType Engine Probe
    '
    '   Ensures that every MovementType defined in tblMovementTypes
    '   and marked Defined? = TRUE has a corresponding implementation
    '   in the MovementType engine (ApplyMovementType).
    '
    '   Any missing MovementType implementation is treated as a
    '   configuration-level fatal error and aggregated into
    '   fatalErrors, so the workflow aborts cleanly before any
    '   Unique Identifiers are processed.
    ' --------------------------------------------------------
    Dim movementEngineError As String
    movementEngineError = ValidateMovementTypeEngine(movementTypeDict)
    
    If movementEngineError <> "" Then
        fatalErrors.Add movementEngineError
    End If
    
    ' --------------------------------------------------------
    ' Validate WriteMode engine
    ' --------------------------------------------------------
    Dim writeModeEngineError As String
    writeModeEngineError = ValidateWriteModeEngine(writeModeDict)
    If writeModeEngineError <> "" Then
        fatalErrors.Add writeModeEngineError
    End If
    
    
    ' --------------------------------------------------------
    ' Validate ConditionOperator engine
    ' --------------------------------------------------------
    Dim operatorEngineError As String
    operatorEngineError = ValidateConditionOperatorEngine(operatorDict)
    If operatorEngineError <> "" Then
        fatalErrors.Add operatorEngineError
    End If
    
    
    ' --------------------------------------------------------
    ' Collect and sort mapping rows (by ExecutionOrder)
    ' --------------------------------------------------------
    Set mappingRows = CollectMappingRows(tblMap, SourceTableFilter, tblMovementLog)
    If mappingRows.Count = 0 Then
        MsgBox "No active mapping rows found for this run.", vbInformation
        Exit Sub
    End If
    
    
    ' --------------------------------------------------------
    ' Validate table-level configuration
    '
    ' Ensures:
    '   - Every table listed in tblWorkflowTables exists
    '   - Every PrimaryKeyField exists in that table
    '   - Every Source/Destination in tblWorkflowMapping is listed
    '     and EligibleForWorkflow = TRUE
    ' --------------------------------------------------------
    Dim tName As Variant
    Dim pkField As String
    Dim tbl As ListObject

    ' Validate each row in tblWorkflowTables
    For Each tName In tablePKDict.Keys
        pkField = tablePKDict(tName)

        Set tbl = GetTableByName(tName)
        If tbl Is Nothing Then
            fatalErrors.Add "Table '" & tName & "' listed in tblWorkflowTables does not exist."
        Else
            If Not HeaderExistsInTable(tbl, pkField) Then
                fatalErrors.Add "PrimaryKeyField '" & pkField & "' not found in table '" & tName & "'."
            End If
        End If
    Next tName

    ' Validate that every table referenced in mapping is eligible
    For Each mappingRow In mappingRows
        sourceTableName = Trim(CStr(mappingRow.Range.Columns(1).Value))
        destTableName = Trim(CStr(mappingRow.Range.Columns(2).Value))

        If Not tableEligibleDict.Exists(sourceTableName) Then
            fatalErrors.Add "Source table '" & sourceTableName & "' is not listed in tblWorkflowTables."
        ElseIf tableEligibleDict(sourceTableName) = False Then
            fatalErrors.Add "Source table '" & sourceTableName & "' is marked as not eligible for workflow."
        End If

        If Not tableEligibleDict.Exists(destTableName) Then
            fatalErrors.Add "Destination table '" & destTableName & "' is not listed in tblWorkflowTables."
        ElseIf tableEligibleDict(destTableName) = False Then
            fatalErrors.Add "Destination table '" & destTableName & "' is marked as not eligible for workflow."
        End If
    Next mappingRow
    
    
    
    ' --------------------------------------------------------
    ' CONFIGURATION VALIDATION (aggregate all fatal errors)
    '
    ' This scans EVERY mapping row before any data movement.
    ' Any configuration error (missing table, missing header,
    ' invalid RuleType, invalid MovementType, invalid WriteMode)
    ' is collected into fatalErrors.
    '
    ' If ANY fatal errors exist, the workflow:
    '   - logs ONE "Aborted" entry
    '   - shows a message box
    '   - exits BEFORE processing any Unique Identifiers
    '
    ' This allows you to fix ALL configuration issues in one go.
    '
    ' Note that only one error per mapping row will be returned
    ' this means that if there are multiple errors in a mapping row
    ' within the configuration table, it only returns the first
    ' However, if there is one error per mapping row it will still
    ' aggregate them
    ' --------------------------------------------------------

        
        For Each mappingRow In mappingRows
        
            ' Validate this mapping row
            configError = ValidateConfigurationForMappingRow(mappingRow, ruleDict, movementTypeDict, operatorDict, writeModeDict)
        
            ' If an error exists, store it for aggregation
            If configError <> "" Then
                fatalErrors.Add configError
            End If
        
        Next mappingRow
        
    ' --------------------------------------------------------
    ' If ANY fatal errors were found, abort the workflow
    ' --------------------------------------------------------
    If fatalErrors.Count > 0 Then
        
            ' Join all error messages into one readable string
            Dim combined As String
            combined = JoinCollection(fatalErrors, " | ")
        
            ' Log ONE aborted entry.
            ' Configuration errors are global, so:
            '   - SourceTable = "Workflow" (label)
            '   - DestinationTable = ""
            '   - ID = ""
            '   - MovementType = ""
            LogMovement tblMovementLog, "Workflow", "", "", "", "Aborted", combined
        
            ' Notify the user
            MsgBox "Workflow aborted due to configuration errors:" & vbCrLf & combined, vbCritical
        
            Exit Sub
        End If
    
    ' ============================================================
    ' TABLE-LEVEL EXECUTION ORDER
    '
    ' We sort the source tables using the declarative
    ' TableExecutionOrder column from tblWorkflowTables.
    '
    ' IMPORTANT:
    '   - This does NOT change row-level ExecutionOrder.
    '   - This does NOT change grouping logic.
    '   - This does NOT change any movement or validation logic.
    '   - Tables with blank execution order run LAST (999999).
    '
    ' This replaces the old dictionary iteration, which was
    ' unpredictable and depended on internal key order.
    ' ============================================================
    
    Set groupedMappings = GroupMappingRowsBySource(mappingRows)
    
    ' RowsToDelete: run-local collection of source rows that moved successfully and should be deleted.
    ' ExecuteMovementForID will add candidates to this collection after a successful transactional move.
    ' RunWorkflow will delete candidates per source table after each group completes.
    Dim RowsToDelete As Collection
    Set RowsToDelete = New Collection
    
    Dim sourceKeys() As Variant
    Dim i As Long, j As Long
    Dim tmpKey As Variant
    
    ' ------------------------------------------------------------
    ' 1. Collect all source table names into an array
    ' ------------------------------------------------------------
    ReDim sourceKeys(1 To groupedMappings.Count)
    i = 1
    For Each groupKey In groupedMappings.Keys
        sourceKeys(i) = groupKey
        i = i + 1
    Next groupKey
    
    ' ------------------------------------------------------------
    ' 2. Sort the source tables by TableExecutionOrder
    '
    '    tableExecOrderDict(tableName) was built earlier from
    '    tblWorkflowTables. Blank values were assigned 999999.
    '
    '    This bubble sort mirrors your existing style in
    '    SortMappingRowsByExecutionOrder — consistent and simple.
    ' ------------------------------------------------------------
    For i = 1 To UBound(sourceKeys) - 1
        For j = i + 1 To UBound(sourceKeys)
            If CLng(tableExecOrderDict(CStr(sourceKeys(i)))) > _
               CLng(tableExecOrderDict(CStr(sourceKeys(j)))) Then
    
                ' Swap the two table names
                tmpKey = sourceKeys(i)
                sourceKeys(i) = sourceKeys(j)
                sourceKeys(j) = tmpKey
            End If
        Next j
    Next i
    

    ' ------------------------------------------------------------
    ' 3. Process each source table in the new sorted order
    '    Sort each group's mapping rows by ExecutionOrder (local),
    '    ensuring primary-key mapping row(s) run first.
    ' ------------------------------------------------------------
    For i = 1 To UBound(sourceKeys)
        groupKey = CStr(sourceKeys(i))
        Set groupRows = groupedMappings(groupKey)
    
        ' Resolve primary key field for this source table (local to this loop)
        Dim groupPKField As String
        groupPKField = ""
    
        ' Defensive type check for the builder output
        If LCase(TypeName(tablePKDict)) <> "scripting.dictionary" And LCase(TypeName(tablePKDict)) <> "dictionary" Then
            Err.Raise vbObjectError + 521, "RunWorkflow", _
                "Unexpected type for tablePKDict: " & TypeName(tablePKDict) & ". Aborting workflow."
        End If
    
        ' Lookup primary key field (try exact key then lowercase key)
        If tablePKDict.Exists(groupKey) Then
            groupPKField = tablePKDict(groupKey)
        ElseIf tablePKDict.Exists(LCase(groupKey)) Then
            groupPKField = tablePKDict(LCase(groupKey))
        Else
            Err.Raise vbObjectError + 520, "RunWorkflow", _
                "PrimaryKeyField not defined for source table '" & groupKey & "'. Aborting workflow."
        End If
    
        ' Sort the group's mapping rows by ExecutionOrder with PK priority
        If groupRows.Count > 1 Then
            SortMappingRowsByExecutionOrderWithPK groupRows, groupPKField
        End If
    
        ' Execute movement for this source table (transactional per-ID)
        ExecuteMovementForID groupRows, ruleDict, tblMovementLog, tablePKDict, RowsToDelete
        
        ' IMPORTANT, THIS IS WHERE THE DELETE OCCURS,
        ' After movement for this source table completes, delete candidates for this source table
        
        Call DeleteRowsForSourceTable(groupKey, RowsToDelete, tblMovementLog, tablePKDict)
    Next i

MsgBox "Workflow movement complete.", vbInformation
Exit Sub

ErrHandler:
    MsgBox "Unexpected error in RunWorkflow: " & Err.Description, vbCritical
End Sub

' ============================================================
' CollectMappingRows
'   Returns a Collection of active mapping rows, optionally
'   filtered by SourceTable. ExecutionOrder is normalised:
'   numeric -> CLng, blank/non-numeric -> 999999.
' ============================================================
Private Function CollectMappingRows(tblMap As ListObject, _
                                    ByVal SourceTableFilter As String, _
                                    tblMovementLog As ListObject) As Collection
    Dim rows As New Collection
    Dim r As ListRow
    Dim activeFlag As Boolean
    Dim srcTable As String
    Dim execOrderRaw As Variant
    Dim execOrderNum As Long

    ' Dictionary to collect normalised rows per source table
    Dim normDict As Object
    Set normDict = CreateObject("Scripting.Dictionary")
    normDict.CompareMode = vbTextCompare

    For Each r In tblMap.ListRows
        activeFlag = CBool(r.Range.Columns(8).Value) ' Active (col 8)
        srcTable = Trim(CStr(r.Range.Columns(1).Value)) ' SourceTable (col 1)
        execOrderRaw = r.Range.Columns(9).Value

        If activeFlag Then
            If SourceTableFilter = "" Or LCase(srcTable) = LCase(SourceTableFilter) Then
                ' Normalize ExecutionOrder: numeric -> CLng, blank/non-numeric -> 999999
                execOrderRaw = r.Range.Columns(9).Value

                If IsError(execOrderRaw) Then
                    execOrderNum = 999999
                ElseIf Trim$(CStr(execOrderRaw)) = "" Then
                    execOrderNum = 999999
                ElseIf Not IsNumeric(execOrderRaw) Then
                    execOrderNum = 999999
                Else
                    execOrderNum = CLng(execOrderRaw)
                End If

                ' Persist normalized value back to the mapping table as a number
                r.Range.Columns(9).Value = execOrderNum

                ' If we normalised (original not numeric or blank), record it
                If Not IsNumeric(execOrderRaw) Or Trim(CStr(execOrderRaw)) = "" Then
                    Dim shortDesc As String
                    ' DestinationTable = col 2, DestinationHeader = col 6, original value shown
                    shortDesc = "Dest=" & Trim(CStr(r.Range.Columns(2).Value)) & _
                                ";DestHeader=" & Trim(CStr(r.Range.Columns(6).Value)) & _
                                ";OrigExec=" & """" & Trim(CStr(execOrderRaw)) & """"

                    If Not normDict.Exists(srcTable) Then
                        normDict(srcTable) = shortDesc
                    Else
                        normDict(srcTable) = normDict(srcTable) & " ; " & shortDesc
                    End If
                End If

                rows.Add r
            End If
        End If
    Next r

    ' Emit one log entry per source table that had normalisations (only if any)
    If normDict.Count > 0 Then
        Dim s As Variant
        For Each s In normDict.Keys
            Dim details As String
            details = "ExecutionOrder normalised for mapping rows: " & normDict(s)
            ' LogMovement(tblMovementLog, SourceTable, DestTable, ID, MovementType, Result, Details)
            LogMovement tblMovementLog, CStr(s), "", "", "N/A", "ExecutionOrder Normalised", details
        Next s
    End If

    Set CollectMappingRows = rows
End Function

' ============================================================
' GroupMappingRowsBySource
'
'   Groups mapping rows ONLY by SourceTable.
'
'   This ensures that ALL destinations for a given source table
'   are processed TOGETHER inside ExecuteMovementForID,
'   enabling transactional (all-or-nothing) movement.
'
'   OUTPUT:
'       Dictionary where:
'           key   = LCase(SourceTable)
'           item  = Collection of mapping rows for that source
' ============================================================
Private Function GroupMappingRowsBySource(mappingRows As Collection) As Object
    Dim groups As Object
    Dim r As ListRow
    Dim src As String
    Dim key As String
    Dim col As Collection
    
    Set groups = CreateObject("Scripting.Dictionary")
    
    For Each r In mappingRows
        src = Trim(CStr(r.Range.Columns(1).Value)) ' SourceTable
        key = LCase(src)
        
        If Not groups.Exists(key) Then
            Set col = New Collection
            groups.Add key, col
        End If
        
        groups(key).Add r
    Next r
    
    Set GroupMappingRowsBySource = groups
End Function

' ============================================================
' SortMappingRowsByExecutionOrderWithPK
'   Sorts a Collection of mapping ListRow objects for a single
'   source table so that:
'     1) mapping row(s) whose SourceHeader (col 5) equals the
'        primaryKeyField are placed first, and
'     2) remaining rows are ordered by ExecutionOrder (col 9),
'        with non-numeric or blank values treated as 999999.
'   Tie-breakers: DestinationTable (col 2) then DestinationHeader (col 6).
' ============================================================
Private Sub SortMappingRowsByExecutionOrderWithPK(rows As Collection, ByVal primaryKeyField As String)
    Dim n As Long, i As Long, j As Long
    Dim arr() As Variant            ' array to hold ListRow object references
    Dim temp As Variant            ' temporary holder for object swaps
    Dim keyI As Long, keyJ As Long ' numeric ExecutionOrder values
    Dim destI As String, destJ As String
    Dim srcFieldI As String, srcFieldJ As String
    Dim destFieldI As String, destFieldJ As String
    Dim pkLower As String

    ' Quick exits for invalid input
    If rows Is Nothing Then Exit Sub
    n = rows.Count
    If n <= 1 Then Exit Sub

    ' Normalize primary key field for case-insensitive comparison
    pkLower = LCase(Trim(primaryKeyField))

    ' Copy collection items into an array using Set to preserve object refs
    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = rows(i)   ' assign ListRow object reference
    Next i

    ' Bubble sort the array using the comparison rules:
    For i = 1 To n - 1
        For j = i + 1 To n
            ' ExecutionOrder safe conversion: non-numeric or blank -> 999999
            On Error Resume Next
            keyI = CLng(arr(i).Range.Columns(9).Value)
            If Err.Number <> 0 Then
                keyI = 999999
                Err.Clear
            End If
            keyJ = CLng(arr(j).Range.Columns(9).Value)
            If Err.Number <> 0 Then
                keyJ = 999999
                Err.Clear
            End If
            On Error GoTo 0

            ' Read SourceHeader for PK membership and DestinationHeader for tie-breakers
            srcFieldI = LCase(Trim(CStr(arr(i).Range.Columns(5).Value)))
            srcFieldJ = LCase(Trim(CStr(arr(j).Range.Columns(5).Value)))

            destFieldI = LCase(Trim(CStr(arr(i).Range.Columns(6).Value)))
            destFieldJ = LCase(Trim(CStr(arr(j).Range.Columns(6).Value)))

            Dim doSwap As Boolean
            doSwap = False

            ' PK rows must come first (compare against SourceHeader)
            If srcFieldI = pkLower And srcFieldJ <> pkLower Then
                doSwap = False
            ElseIf srcFieldJ = pkLower And srcFieldI <> pkLower Then
                doSwap = True
            Else
                ' Neither or both are PK: compare ExecutionOrder
                If keyI > keyJ Then
                    doSwap = True
                ElseIf keyI = keyJ Then
                    ' Deterministic tie-breaker: DestinationTable then DestinationHeader
                    destI = LCase(Trim(CStr(arr(i).Range.Columns(2).Value)))
                    destJ = LCase(Trim(CStr(arr(j).Range.Columns(2).Value)))
                    If destI > destJ Then
                        doSwap = True
                    ElseIf destI = destJ Then
                        If destFieldI > destFieldJ Then
                            doSwap = True
                        End If
                    End If
                End If
            End If

            ' Perform object swap using Set to preserve references
            If doSwap Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i

    ' Rebuild the original collection in the sorted order:
    For i = rows.Count To 1 Step -1
        rows.Remove i
    Next i

    For i = 1 To n
        rows.Add arr(i)
    Next i
End Sub

' ============================================================
' ExecuteMovementForID  (Transactional Workflow Engine)
'
' PURPOSE:
'   Executes workflow movement for a single Unique Identifier across one
'   or more destination tables, using a fully transactional
'   (all-or-nothing) model.
'
'       • All INCLUDED destinations must validate successfully
'         before ANY movement occurs.
'
'       • If ANY included destination fails validation, then:
'             - NO destinations receive data
'             - Each included destination logs "Not moved"
'             - The source row remains untouched
'
'       • If ALL included destinations pass validation, then:
'             - ALL included destinations receive data
'             - Each included destination logs "Moved"
'
'
' DESTINATION ROUTING (Condition-Driven):
'   A destination participates in the transaction ONLY if at
'   least one of its mapping rows has a condition that evaluates
'   TRUE for the current Unique Identifier (primary key).
'
'   This routing logic is fully table-driven. No column names or
'   workflow rules are hard-coded in the VBA.
'
'
' VALIDATION PHASE:
'   For each INCLUDED destination:
'       • All mapping rows whose conditions are met are validated
'         using their RuleType (NotBlank, Number, Date, etc.).
'
'       • Validation errors are collected PER DESTINATION.
'
'   If ANY included destination has validation errors:
'       - The transaction fails.
'       - No movement occurs.
'       - Each included destination logs "Not moved" with its
'         own error list.
'
'
' MOVEMENT PHASE:
'   Only executed if ALL included destinations pass validation.
'
'   For each included destination:
'       • All applicable mapping rows perform movement
'         (Create/Append) via PerformMovementForRow.
'
'       • Duplicate detection still applies per destination.
'
'   Each included destination logs:
'       • "Moved"  (success)
'       • "Duplicate" (if Unique Identifier (primary key value) already exists)
'
'
' LOGGING:
'   Logging remains PER DESTINATION.
'
'   This means:
'       • DestinationTable1 gets its own log entry.
'       • DestinationTable2 gets its own log entry.
'
'   Even though movement is transactional, logging is still
'   destination-specific for clarity and consistency.
'
'
' KEY BENEFITS OF THIS MODEL:
'   • No partial movement — destinations never get out of sync.
'   • Clean delete logic — a source row is only eligible for
'     deletion once ALL included destinations have moved.
'   • Fully declarative routing — controlled entirely by the
'     mapping table (ConditionField/Operator/Value).
'   • No hard-coded workflow rules.
'   • Predictable, auditable behaviour.
'
' ============================================================
Public Sub ExecuteMovementForID(mappingRows As Collection, _
                                ruleDict As Object, _
                                tblMovementLog As ListObject, _
                                tablePKDict As Object, _
                                RowsToDelete As Collection)

    Dim firstMapRow As ListRow
    Dim sourceTableName As String
    Dim movementType As String
    
    Dim tblSource As ListObject
    Dim srcDataRow As ListRow
    Dim pkValue As String    ' Primary key value for this record
    
    Dim primaryKeyField As String
    
    Dim destGroups As Object          ' Dictionary: key = DestinationTable, value = Collection of mapping rows
    Dim destName As Variant
    Dim destRows As Collection
    
    Dim includedDestinations As Object ' Dictionary of destinations that participate
    Dim validationErrors As Object     ' Dictionary: key = DestinationTable, value = error string
    
    Dim tblDest As ListObject
    Dim mapRow As ListRow
    Dim errText As String
    Dim resultText As String
    
    Dim allValid As Boolean
    Dim overallResult As String
    Dim overallDetails As String
    
    ' --------------------------------------------------------
    ' Resolve SourceTable and MovementType from first row
    ' --------------------------------------------------------
    Set firstMapRow = mappingRows(1)
    sourceTableName = Trim(CStr(firstMapRow.Range.Columns(1).Value))
    movementType = Trim(CStr(firstMapRow.Range.Columns(3).Value))
    
    ' Resolve primary key for the source table
    primaryKeyField = tablePKDict(sourceTableName)
    If primaryKeyField = "" Then
        Err.Raise vbObjectError + 513, "RunWorkflow", _
            "Primary key not defined for table '" & sourceTableName & "'. " & _
            "Validation should have prevented this."
    End If
    
    Set tblSource = GetTableByName(sourceTableName)
    If tblSource Is Nothing Then Exit Sub
    
    ' --------------------------------------------------------
    ' Group mapping rows by DestinationTable
    ' --------------------------------------------------------
    Set destGroups = CreateObject("Scripting.Dictionary")
    
    For Each mapRow In mappingRows
        Dim dest As String
        dest = Trim(CStr(mapRow.Range.Columns(2).Value)) ' DestinationTable
        
        If Not destGroups.Exists(dest) Then
            destGroups.Add dest, New Collection
        End If
        
        destGroups(dest).Add mapRow
    Next mapRow
    
    ' --------------------------------------------------------
    ' LOOP: For each Unique Identifier (primary key value) in the source table
    ' --------------------------------------------------------
    For Each srcDataRow In tblSource.ListRows
        
        pkValue = CStr(GetCellByHeader(tblSource, srcDataRow.Range, primaryKeyField))
        If Len(Trim$(pkValue)) = 0 Then GoTo NextIssue
        
        ' ====================================================
        ' 1. ROUTING PHASE — determine INCLUDED destinations
        ' ====================================================
        Set includedDestinations = CreateObject("Scripting.Dictionary")
        
        For Each destName In destGroups.Keys
            
            Set destRows = destGroups(destName)
            
            Dim atLeastOneRowApplies As Boolean
            atLeastOneRowApplies = False
            
            ' Check conditions for each mapping row
            For Each mapRow In destRows
                If ConditionIsMet(tblSource, srcDataRow, mapRow) Then
                    atLeastOneRowApplies = True
                    Exit For
                End If
            Next mapRow
            
            ' If at least one mapping row applies - include destination
            If atLeastOneRowApplies Then
                includedDestinations.Add destName, destRows
            Else
                ' Log destinations where NO mapping rows applied for this Unique Identifier (primary key value)
                ' This covers cases where:
                '   - All mapping rows have conditions evaluated FALSE (e.g. Status blank when expecting Complete/Cancelled)
                '   - All mapping rows used unsupported operators
                ' It only runs when the destination is NOT included and would otherwise be completely silent.
                overallResult = "Not moved"
                overallDetails = "Not moved — no mapping rows applied. For this destination, every mapping-row condition evaluated FALSE or used an unsupported operator."
                
                LogMovement tblMovementLog, sourceTableName, destName, pkValue, movementType, overallResult, overallDetails
                
            
            End If
        Next destName
        
        
        ' ============================================================
        ' MASTER-CONDITION RULE FOR Primary Key
        '
        ' PURPOSE:
        '   Prevents partial movement when the Primary Key mapping row
        '   is ineligible. If the Primary Key's row's condition evaluates
        '   FALSE for a destination, the entire destination must be
        '   excluded — even if other rows have TRUE conditions.
        '
        ' WHY:
        '   pkValue is the primary key for all destinations. Without
        '   it, movement would create orphaned rows, blank-ID rows,
        '   or overwritten rows on the next run.
        '
        ' BEHAVIOUR:
        '   For each included destination:
        '       • Find the mapping row where FieldName = pkvalue
        '         (case-insensitive, trimmed).
        '       • Evaluate its condition.
        '       • If FALSE - remove destination from included list.
        '       • Log a clear message explaining why it was excluded.
        '
        ' This is a LOCAL rule:
        '   Only destinations that actually have a pkValue row are
        '   subject to this master-condition behaviour.
        ' ============================================================
        
        Dim destToRemove As Collection
        Set destToRemove = New Collection
        
        Dim haspkRow As Boolean
        Dim pkConditionMet As Boolean
        Dim fieldName As String
        
        For Each destName In includedDestinations.Keys
            
            haspkRow = False
            pkConditionMet = False
            
            Set destRows = includedDestinations(destName)
            
            ' Find pk (primary key/ unique identifier) row for this destination
            For Each mapRow In destRows
                fieldName = LCase(Trim(CStr(mapRow.Range.Columns(4).Value))) ' FieldName column
                
                ' use declarative primary key instead of hard-coded "id"
                If LCase(fieldName) = LCase(primaryKeyField) Then
                    haspkRow = True
                    
                    ' Evaluate the condition for the Unique Identifier (primary key) row
                    If ConditionIsMet(tblSource, srcDataRow, mapRow) Then
                        pkConditionMet = True
                    End If
                    
                    Exit For
                End If
            Next mapRow
            
            ' If destination has a pk row row AND its condition failed - mark for removal
            If haspkRow And Not pkConditionMet Then
                destToRemove.Add destName
                
                ' Log the reason immediately (destination-specific)
                LogMovement tblMovementLog, sourceTableName, destName, pkValue, movementType, _
                            "Excluded - No movement", _
                            "Master condition failed: Primary Key mapping row was ineligible for this Unique identifier (Primary key, destination excluded)."
            End If
        Next destName
        
        ' Remove destinations whose unique identifier (primary key) condition failed
        For Each destName In destToRemove
            includedDestinations.Remove destName
        Next destName
        
        ' ============================================================
        ' END OF MASTER-CONDITION BLOCK
        ' ============================================================
        
        
        
        ' If no destinations included, nothing to do
        If includedDestinations.Count = 0 Then GoTo NextIssue
        
        ' ====================================================
        ' 2. VALIDATION PHASE — validate all included dests
        ' ====================================================
        Set validationErrors = CreateObject("Scripting.Dictionary")
        allValid = True
        
        For Each destName In includedDestinations.Keys
            
            Set destRows = includedDestinations(destName)
            Set tblDest = GetTableByName(destName)
            
            Dim combinedErrors As String
            combinedErrors = ""
            
            ' Validate each mapping row for this destination
            For Each mapRow In destRows
                
                If ConditionIsMet(tblSource, srcDataRow, mapRow) Then
                    errText = ValidateRowForMappingRow(tblSource, tblDest, srcDataRow, mapRow, ruleDict)
                    
                    If errText <> "" Then
                        If combinedErrors = "" Then
                            combinedErrors = errText
                        Else
                            combinedErrors = combinedErrors & " | " & errText
                        End If
                    End If
                End If
            Next mapRow
            
            validationErrors.Add destName, combinedErrors
            
            If combinedErrors <> "" Then
                allValid = False
            End If
        Next destName
        
        ' ====================================================
        ' 3. DECISION — if ANY included destination invalid:
        '       - NO MOVEMENT
        '       - Log Not moved for each included destination
        ' ====================================================
        If Not allValid Then
            
            For Each destName In includedDestinations.Keys
                overallResult = "Not moved"
                overallDetails = validationErrors(destName)
                
                LogMovement tblMovementLog, sourceTableName, destName, pkValue, movementType, overallResult, overallDetails
            Next destName
            
            GoTo NextIssue
        End If
        
        ' ====================================================
        ' 4. MOVEMENT PHASE — all included destinations valid
        ' ====================================================
        ' aggregate success flag for all included destinations
        Dim allSucceeded As Boolean
        allSucceeded = True    ' assume success until a destination fails
        
        For Each destName In includedDestinations.Keys
        
            Set destRows = includedDestinations(destName)
            Set tblDest = GetTableByName(destName)
        
            ' Initialize per-destination result
            overallResult = "Moved"
            overallDetails = ""
        
            For Each mapRow In destRows
        
                If ConditionIsMet(tblSource, srcDataRow, mapRow) Then
        
                    ' Perform movement for this mapping row and capture the result
                    resultText = PerformMovementForRow(tblSource, tblDest, srcDataRow, mapRow, pkValue, primaryKeyField)
        
                    ' propagate resultText into overallResult deterministically ---
                    ' This ensures failures returned by PerformMovementForRow are not ignored.
                    If LCase$(Trim$(resultText)) = "duplicate" Then
                        overallResult = "Duplicate"
                        overallDetails = "Unique ID '" & pkValue & "' already exists in destination table '" & destName & "'."
                        ' Duplicate treated as terminal success for this destination
                        Exit For
                    ElseIf LCase$(Trim$(resultText)) = "moved" Then
                        overallResult = "Moved"
                        ' continue checking other mapping rows for this destination
                    Else
                        ' Any other result (validation failure, "Not moved", error text) is a failure for this destination
                        overallResult = resultText
                        If overallDetails = "" Then
                            overallDetails = resultText
                        Else
                            overallDetails = overallDetails & " ; " & resultText
                        End If
                        ' Mark aggregate as failed and stop processing this destination
                        allSucceeded = False    ' update aggregate flag
                        Exit For
                    End If
                    ' --- END propagation logic ---
        
                End If
            Next mapRow
        
            ' Log the per-destination result (unchanged behaviour)
            LogMovement tblMovementLog, sourceTableName, destName, pkValue, movementType, overallResult, overallDetails
        
            ' --- Update last moved marker for this Source->Dest when the move succeeded ---
            If Not UpdateLastMovedTable(sourceTableName, destName, pkValue) Then
                Call LogMovement(tblMovementLog, sourceTableName, destName, pkValue, "MetaUpdate", "Failed", "tblLastMoved not found or update failed")
            End If
        
            '  ensure we also treat explicit non-success results as aggregate failure ---
            ' (This is defensive: if PerformMovementForRow returned "Not moved" or other text,
            '  allSucceeded was already set to False above. This double-check keeps behaviour explicit.)
            Dim tmpRes As String
            tmpRes = LCase$(Trim$(overallResult))
            If Not (tmpRes = "moved" Or tmpRes = "duplicate") Then
                allSucceeded = False
            End If
            ' --- END aggregate update ---
        
        Next destName
        
        ' --- Deterministic delete decision (replace fragile last-result check) ---
        ' Queue source row for deletion only if:
        '   1) At least one destination was included for this PK, and
        '   2) Every included destination ended in an acceptable state.
        ' Acceptable states here are "Moved" and "Duplicate" (treat duplicates as success).
        If includedDestinations.Count > 0 And allSucceeded Then
            Dim candidate As Object
            Set candidate = CreateObject("Scripting.Dictionary")
            candidate("SourceTable") = tblSource.Name
            candidate("SourcePK") = CStr(pkValue)
            candidate("SourceRowIndex") = srcDataRow.Index
        
            ' Avoid duplicate queue entries
            Dim alreadyQueued As Boolean
            alreadyQueued = False
            Dim q As Variant
            For Each q In RowsToDelete
                If q("SourceTable") = candidate("SourceTable") And q("SourcePK") = candidate("SourcePK") Then
                    alreadyQueued = True
                    Exit For
                End If
            Next q
        
            If Not alreadyQueued Then RowsToDelete.Add candidate
        
            ' Optional: log that the row was queued for deletion (helps auditing)
            LogMovement tblMovementLog, tblSource.Name, "", CStr(pkValue), "N/A", "Queued for Delete", _
                        "All included destinations succeeded; source primary key row eligible for removal"
        Else
            ' Optional: log why not queued (useful for debugging)
            ' LogMovement tblMovementLog, tblSource.Name, "", CStr(pkValue), "N/A", "Not queued for Delete", _
            '            "IncludedCount=" & includedDestinations.Count & "; AllSucceeded=" & CStr(allSucceeded)
        End If
        ' --- END deterministic delete decision ---

        
NextIssue:
    Next srcDataRow
End Sub

' ============================================================
' DeleteRowsForSourceTable
'   Deletes all RowsToDelete candidates for a single source
'   table. Prefers delete-by-PK, falls back to delete-by-index.
'   Logs failures via LogMovement and removes processed items
'   from the RowsToDelete collection.
' ============================================================
Private Sub DeleteRowsForSourceTable(ByVal sourceTableName As String, _
                                     RowsToDelete As Collection, _
                                     tblMovementLog As ListObject, _
                                     tablePKDict As Object)

    Dim toDelete As New Collection
    Dim i As Long, a As Long, b As Long
    Dim tmp As Variant
    Dim item As Variant

    ' 1) Collect candidates for this source table
    For i = 1 To RowsToDelete.Count
        Set item = RowsToDelete(i)
        If LCase(CStr(item("SourceTable"))) = LCase(sourceTableName) Then
            toDelete.Add item
        End If
    Next i

    ' Nothing to do
    If toDelete.Count = 0 Then Exit Sub

    ' 2) Sort by SourceRowIndex descending to avoid index shifting issues
    If toDelete.Count > 1 Then
        Dim n As Long
        n = toDelete.Count
    
        ' Copy collection items into an array to allow swapping
        Dim arr() As Variant
        ReDim arr(1 To n)
        For a = 1 To n
            Set arr(a) = toDelete(a)
        Next a
    
        ' Bubble sort the array by SourceRowIndex descending
        For a = 1 To n - 1
            For b = a + 1 To n
                If CLng(arr(a)("SourceRowIndex")) < CLng(arr(b)("SourceRowIndex")) Then
                    Set tmp = arr(a)
                    Set arr(a) = arr(b)
                    Set arr(b) = tmp
                End If
            Next b
        Next a
    
        ' Rebuild the collection in sorted order
        For a = toDelete.Count To 1 Step -1
            toDelete.Remove a
        Next a
        For a = 1 To n
            toDelete.Add arr(a)
        Next a
    End If

    ' 3) Attempt deletes and log failures
    Dim deleteSucceeded As Boolean
    For i = 1 To toDelete.Count
        Set item = toDelete(i)
        deleteSucceeded = False
    
        ' Prefer delete by PK if available
        If Len(Trim$(CStr(item("SourcePK")))) > 0 Then
            deleteSucceeded = DeleteSourceRowByPK(CStr(item("SourceTable")), CStr(item("SourcePK")), tablePKDict)
            If deleteSucceeded Then
                ' --- confirm actual deletion and which method was used (DeleteSourceRowByPK) ---
                Call LogMovement(tblMovementLog, CStr(item("SourceTable")), "", CStr(item("SourcePK")), "Delete", "Deleted", _
                                 "Deleted by primary key using DeleteSourceRowByPK.")
            End If
        End If
    
        ' Fallback: delete by index
        If Not deleteSucceeded Then
            On Error Resume Next
            deleteSucceeded = DeleteSourceRowByIndex(CStr(item("SourceTable")), CLng(item("SourceRowIndex")))
            If Err.Number <> 0 Then
                ' Optional: capture unexpected error text
                Dim delErr As String
                delErr = "Error " & Err.Number & ": " & Err.Description
                Err.Clear
            Else
                delErr = ""
            End If
            On Error GoTo 0
    
            If deleteSucceeded Then
                ' --- confirm actual deletion and which method was used (DeleteSourceRowByIndex) ---
                Call LogMovement(tblMovementLog, CStr(item("SourceTable")), "", CStr(item("SourcePK")), "Delete", "Deleted", _
                                 "Deleted by row index using DeleteSourceRowByIndex.")
            End If
        End If
    
        ' If delete failed, log a failure row so it is visible for manual remediation
        If Not deleteSucceeded Then
            Dim failDetails As String
            If Len(Trim$(delErr)) > 0 Then
                failDetails = "Auto-delete failed; " & delErr
            Else
                failDetails = "Auto-delete failed; neither PK nor index deletion succeeded."
            End If
    
            Call LogMovement(tblMovementLog, CStr(item("SourceTable")), "", CStr(item("SourcePK")), "Delete", "Delete Failed", failDetails)
        End If
    Next i
End Sub


' ============================================================
' DeleteSourceRowByPK
'   Finds the table, resolves its PK column from tablePKDict,
'   locates the row by exact match and deletes it.
'   Returns True on success, False otherwise.
' ============================================================
Private Function DeleteSourceRowByPK(ByVal tableName As String, ByVal pkValue As String, tablePKDict As Object) As Boolean
    Dim tbl As ListObject
    Dim pkField As String
    Dim pkCol As ListColumn
    Dim rng As Range

    DeleteSourceRowByPK = False

    Set tbl = GetTableByName(tableName)
    If tbl Is Nothing Then Exit Function

    On Error Resume Next
    pkField = tablePKDict(tableName)
    If pkField = "" Then pkField = tablePKDict(LCase(tableName))
    On Error GoTo 0
    If Trim$(pkField) = "" Then Exit Function

    On Error Resume Next
    Set pkCol = tbl.ListColumns(pkField)
    On Error GoTo 0
    If pkCol Is Nothing Then Exit Function

    If pkCol.DataBodyRange Is Nothing Then Exit Function

    On Error Resume Next
    Set rng = pkCol.DataBodyRange.Find(What:=pkValue, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function

    ' Delete the found DataBody row
    tbl.ListRows(rng.row - tbl.DataBodyRange.row + 1).Delete
    DeleteSourceRowByPK = True
End Function


' ============================================================
' DeleteSourceRowByIndex
'   Deletes a DataBody row by its 1-based index within the table.
'   Returns True on success, False otherwise.
' ============================================================
Private Function DeleteSourceRowByIndex(ByVal tableName As String, ByVal rowIndex As Long) As Boolean
    Dim tbl As ListObject
    DeleteSourceRowByIndex = False

    Set tbl = GetTableByName(tableName)
    If tbl Is Nothing Then Exit Function

    If tbl.ListRows.Count = 0 Then Exit Function
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Function

    tbl.ListRows(rowIndex).Delete
    DeleteSourceRowByIndex = True
End Function

' ============================================================
' ConditionIsMet
'
' PURPOSE:
'   Evaluates the ConditionField / ConditionOperator / ConditionValue
'   for a given source data row.
'
'   Supports a probe mode (tblSource = Nothing) so the validation
'   engine can check whether an operator is implemented.
'
' PARAMETERS:
'   operatorOverride:
'       - When provided, overrides the operator from the mapping row.
'       - Used only by the operator engine probe.
'
'   operatorResult:
'       - Returns "ENGINE_NOT_IMPLEMENTED" during probe mode when
'         the operator is not supported by the engine.
'
' ============================================================
Public Function ConditionIsMet(tblSource As ListObject, _
                               srcDataRow As ListRow, _
                               mapRow As ListRow, _
                               Optional ByVal operatorOverride As String = "", _
                               Optional ByRef operatorResult As String = "") As Boolean
    Dim condField As String
    Dim condOp As String
    Dim condValue As String
    Dim srcValue As String
    Dim headerCell As Range
    Dim colIndex As Long
    Dim isProbe As Boolean
    Dim parts() As String
    Dim p As Variant

    operatorResult = ""
    ConditionIsMet = False

    ' --------------------------------------------------------
    ' Detect probe mode (called from ValidateConditionOperatorEngine)
    ' --------------------------------------------------------
    isProbe = (tblSource Is Nothing)

    ' Resolve operator
    If operatorOverride <> "" Then
        condOp = UCase$(Trim$(operatorOverride))
    ElseIf Not isProbe Then
        condOp = UCase$(Trim$(CStr(mapRow.Range.Columns(11).Value))) ' ConditionOperator
    Else
        condOp = ""
    End If

    ' In probe mode we MUST NOT touch mapRow or tblSource.
    ' We still flow through the Select Case so Case Else can self-report.
    If Not isProbe Then
        ' Read condition field/value from mapping row
        condField = Trim$(CStr(mapRow.Range.Columns(10).Value))   ' ConditionField
        condValue = Trim$(CStr(mapRow.Range.Columns(12).Value))   ' ConditionValue

        ' No condition - always applies
        If condField = "" Then
            ConditionIsMet = True
            Exit Function
        End If

        ' Find the source column by header name
        colIndex = 0
        For Each headerCell In tblSource.HeaderRowRange
            If Trim$(CStr(headerCell.Value)) = condField Then
                colIndex = headerCell.Column - tblSource.Range.Columns(1).Column + 1
                Exit For
            End If
        Next headerCell

        ' Header not found - treat as not met
        If colIndex = 0 Then
            ConditionIsMet = False
            Exit Function
        End If

        srcValue = Trim$(CStr(srcDataRow.Range.Cells(1, colIndex).Value))
    End If

    Select Case condOp

        Case "=", ""
            If isProbe Then Exit Function
            ConditionIsMet = (StrComp(srcValue, condValue, vbTextCompare) = 0)

        Case "<>"
            If isProbe Then Exit Function
            ConditionIsMet = (StrComp(srcValue, condValue, vbTextCompare) <> 0)

        Case "ISBLANK"
            If isProbe Then Exit Function
            ConditionIsMet = (srcValue = "")

        Case "ISNOTBLANK"
            If isProbe Then Exit Function
            ConditionIsMet = (srcValue <> "")

        Case "IN"
            If isProbe Then Exit Function
            parts = Split(condValue, ",")
            For Each p In parts
                If StrComp(srcValue, Trim$(CStr(p)), vbTextCompare) = 0 Then
                    ConditionIsMet = True
                    Exit Function
                End If
            Next p
            ConditionIsMet = False

        Case Else
            ' Unknown operator:
            '   - Real mode: treat as FALSE.
            '   - Probe mode: self-report as not implemented.
            If isProbe Then
                operatorResult = "ENGINE_NOT_IMPLEMENTED"
            End If
            ConditionIsMet = False
    End Select
End Function

