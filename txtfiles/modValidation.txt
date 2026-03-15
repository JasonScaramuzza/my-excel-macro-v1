Attribute VB_Name = "modValidation"

' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.0 (14 March 2026)
'
'  Module: [modValidation]
'  Purpose: [Contains the validation functions used in the WorkFlowEngine]
' ================================================================================


Option Explicit

' ============================================================
' CONFIGURATION VALIDATION
' ============================================================
Public Function ValidateConfigurationForMappingRow(mappingRow As ListRow, _
                                                   ruleDict As Object, _
                                                   movementTypeDict As Object, _
                                                   operatorDict As Object, _
                                                   writeModeDict As Object) As String
    Dim sourceTableName As String
    Dim destTableName As String
    Dim movementType As String
    Dim sourceHeader As String
    Dim destHeader As String
    Dim ruleType As String
    Dim writeMode As String
    Dim conditionOperator As String
    Dim conditionField As String

    
    Dim tblSource As ListObject
    Dim tblDest As ListObject
    
    sourceTableName = Trim(CStr(mappingRow.Range.Columns(1).Value))
    destTableName = Trim(CStr(mappingRow.Range.Columns(2).Value))
    movementType = Trim(CStr(mappingRow.Range.Columns(3).Value))
    sourceHeader = Trim(CStr(mappingRow.Range.Columns(5).Value))
    destHeader = Trim(CStr(mappingRow.Range.Columns(6).Value))
    ruleType = Trim(CStr(mappingRow.Range.Columns(7).Value))
    writeMode = Trim(CStr(mappingRow.Range.Columns(13).Value))
    conditionOperator = Trim(CStr(mappingRow.Range.Columns(11).Value))
    conditionField = Trim(CStr(mappingRow.Range.Columns(10).Value))

    
    ' --------------------------------------------------------
    ' Validate tables exist
    ' --------------------------------------------------------
    Set tblSource = GetTableByName(sourceTableName)
    If tblSource Is Nothing Then
        ValidateConfigurationForMappingRow = "Source table '" & sourceTableName & "' not found."
        Exit Function
    End If
    
    Set tblDest = GetTableByName(destTableName)
    If tblDest Is Nothing Then
        ValidateConfigurationForMappingRow = "Destination table '" & destTableName & "' not found."
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' Validate headers
    ' --------------------------------------------------------
    If sourceHeader <> "" Then
        If Not HeaderExistsInTable(tblSource, sourceHeader) Then
            ValidateConfigurationForMappingRow = "Source header '" & sourceHeader & "' not found in " & sourceTableName
            Exit Function
        End If
    End If
    
    If destHeader <> "" Then
        If Not HeaderExistsInTable(tblDest, destHeader) Then
            ValidateConfigurationForMappingRow = "Destination header '" & destHeader & "' not found in " & destTableName
            Exit Function
        End If
    End If
    
    If conditionField <> "" Then
        If Not HeaderExistsInTable(tblSource, conditionField) Then
            ValidateConfigurationForMappingRow = "ConditionField '" & conditionField & "' not found in source table '" & sourceTableName & "'."
            Exit Function
        End If
    End If

    
    ' --------------------------------------------------------
    ' Validate RuleType (TABLE-DRIVEN)
    ' --------------------------------------------------------
    If ruleType <> "" And ruleType <> "Optional" Then
        If Not ruleDict.Exists(ruleType) Then
            ValidateConfigurationForMappingRow = "RuleType '" & ruleType & "' not listed in tblRuleList."
            Exit Function
        ElseIf ruleDict(ruleType) = False Then
            ValidateConfigurationForMappingRow = "RuleType '" & ruleType & "' is listed but Defined? = FALSE."
            Exit Function
        End If
    End If
    
    ' --------------------------------------------------------
    ' Validate MovementType (TABLE-DRIVEN)
    ' --------------------------------------------------------
    If Not movementTypeDict.Exists(movementType) Then
        ValidateConfigurationForMappingRow = "MovementType '" & movementType & "' not listed in tblMovementTypes."
        Exit Function
    ElseIf movementTypeDict(movementType) = False Then
        ValidateConfigurationForMappingRow = "MovementType '" & movementType & "' is listed but Defined? = FALSE."
        Exit Function
    End If
    
    ' --------------------------------------------------------
    ' Validate Conditional Operator (TABLE-DRIVEN)
    ' --------------------------------------------------------
    If conditionOperator <> "" Then
        If Not operatorDict.Exists(conditionOperator) Then
            ValidateConfigurationForMappingRow = "Operator '" & conditionOperator & "' not listed in tblConditionalOperators."
            Exit Function
        ElseIf operatorDict(conditionOperator) = False Then
            ValidateConfigurationForMappingRow = "Operator '" & conditionOperator & "' is listed but Defined? = FALSE."
            Exit Function
        End If
    End If

    ' --------------------------------------------------------
    ' Validate WriteMode (TABLE-DRIVEN)
    ' --------------------------------------------------------
    If Not writeModeDict.Exists(writeMode) Then
        ValidateConfigurationForMappingRow = "WriteMode '" & writeMode & "' not listed in tblWriteModes."
        Exit Function
    ElseIf writeModeDict(writeMode) = False Then
        ValidateConfigurationForMappingRow = "WriteMode '" & writeMode & "' is listed but Defined? = FALSE."
        Exit Function
    End If

        
    ValidateConfigurationForMappingRow = ""
End Function

' ============================================================
' ValidateRowForMappingRow
'
'   Performs ALL row-level validation for a single Primary Key (unique Identifier)
'   under a single mapping row.
'
'   RESPONSIBILITIES:
'   -----------------------------------------------
'   • Evaluate RuleTypes (NotBlank, MustBeDate, etc.).
'   • Evaluate Conditions (e.g. Status IN (...)).
'   • Check header existence and data availability.
'   • Return a SINGLE concatenated error string containing
'     ALL validation failures for that Unique Identifier.
'
'   RETURN VALUE:
'   -----------------------------------------------
'   • "" (empty string) - row is valid and movement may proceed.
'   • "<error1> | <error2> | <error3> | ..." - row is invalid.
'
'   LOGGING:
'   -----------------------------------------------
'   • This function does NOT log anything.
'     ExecuteMovementForID handles logging based on the
'     returned error string.
'
'   PURPOSE:
'   -----------------------------------------------
'   Centralises all validation logic so that movement
'   functions remain clean and focused on writing data.
' ============================================================
Public Function ValidateRowForMappingRow(tblSource As ListObject, _
                                         tblDest As ListObject, _
                                         srcDataRow As ListRow, _
                                         mappingRow As ListRow, _
                                         ruleDict As Object) As String
    Dim ruleType As String
    Dim sourceHeader As String
    Dim cellValue As Variant
    Dim errMsg As String
    
    ruleType = Trim(CStr(mappingRow.Range.Columns(7).Value))   'RuleType
    sourceHeader = Trim(CStr(mappingRow.Range.Columns(5).Value)) 'SourceHeader
    
    If ruleType = "" Or ruleType = "Optional" Then
        ValidateRowForMappingRow = ""
        Exit Function
    End If
    
    ' At this point, configuration has already ensured ruleType is valid
    cellValue = GetCellByHeader(tblSource, srcDataRow.Range, sourceHeader)
    errMsg = ApplyRule(ruleType, sourceHeader, cellValue)
    
    ValidateRowForMappingRow = errMsg
End Function


' ============================================================
' RULE ENGINE
' ============================================================
Public Function ApplyRule(ruleType As String, fieldName As String, val As Variant) As String
    Dim txt As String
    txt = Trim(CStr(val))
    
    Select Case ruleType
    
        Case "NotBlank"
            If txt = "" Then
                ApplyRule = fieldName & ": NotBlank failed (must not be empty)"
            End If
        
        Case "Optional"
            ' Always passes
        
        Case "Date"
            If Not IsDate(val) Then
                ApplyRule = fieldName & ": Date failed (must be a valid date)"
            End If
        
        Case "Number"
            If Not IsNumeric(val) Then
                ApplyRule = fieldName & ": Number failed (must be numeric)"
            End If
        
        Case "MustBeYes"
            If LCase(Trim(CStr(val))) <> "yes" Then
                ApplyRule = fieldName & ": MustBeYes failed (must be 'Yes')"
            End If
        
        Case "MustBeYesOrNo"
            Dim yn As String
            yn = LCase(Trim(CStr(val)))
            If yn <> "yes" And yn <> "no" Then
                ApplyRule = fieldName & ": MustBeYesOrNo failed (must be Yes or No)"
            End If
        
        Case "MustBeDateOrNA"
            Dim d As String
            d = LCase(Trim(CStr(val)))
            If d = "na" Or d = "n/a" Then Exit Function
            If Not IsDate(val) Then
                ApplyRule = fieldName & ": MustBeDateOrNA failed (must be a date or NA)"
            End If
        
        Case "MustBeCompleteOrCancelled"
            Dim cc As String
            cc = LCase(Trim(CStr(val)))
            If cc <> "complete" And cc <> "cancelled" Then
                ApplyRule = fieldName & ": MustBeCompleteOrCancelled failed (must be Complete or Cancelled)"
            End If
        
        Case "MustBeBlankOrNAOrDate"
            Dim bnd As String
            bnd = LCase(Trim(CStr(val)))
            If bnd = "" Or bnd = "na" Or bnd = "n/a" Then Exit Function
            If Not IsDate(val) Then
                ApplyRule = fieldName & ": MustBeBlankOrNAOrDate failed (must be blank, NA, or a date)"
            End If
        
        Case "MustBeCompleteOrNotRequired"
            Dim cnr As String
            cnr = LCase(Trim(CStr(val)))
            If cnr <> "complete" And cnr <> "not required" Then
                ApplyRule = fieldName & ": MustBeCompleteOrNotRequired failed (must be Complete or Not Required)"
            End If
        Case "MustBeBlankOrDate"
            Dim bd As String
            bd = Trim(CStr(val))
            If bd = "" Then Exit Function
            If Not IsDate(val) Then
                ApplyRule = fieldName & ": MustBeBlankOrDate failed (must be blank or a valid date)"
            End If
                    
        Case Else
            ApplyRule = fieldName & ": RuleType '" & ruleType & "' not implemented."
    End Select
End Function



' ============================================================
' BuildRuleDictionary
' ============================================================
Public Function BuildRuleDictionary(tblRuleList As ListObject) As Object
    Dim dict As Object
    Dim r As ListRow
    Dim ruleType As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each r In tblRuleList.ListRows
        ruleType = Trim(CStr(r.Range.Columns(1).Value))
        If ruleType <> "" Then
            dict(ruleType) = CBool(r.Range.Columns(3).Value)
        End If
    Next r
    
    Set BuildRuleDictionary = dict
End Function


'=========================================================================================
' ValidateRuleEngine
'
' PURPOSE:
'   Ensures that every RuleType defined in the Rule List table is actually implemented
'   inside the VBA rule engine (ApplyRule). This prevents a silent failure scenario where:
'       - A RuleType exists in the configuration table
'       - But the corresponding Case block is missing in ApplyRule
'       - And the mapping-row condition evaluates FALSE, meaning ApplyRule is never called
'
'   Without this check, a missing RuleType implementation can be completely hidden at
'   runtime, because the condition logic may prevent ApplyRule from ever executing.
'
' HOW IT WORKS:
'   - The ruleDict contains every RuleType defined in the Rule List table.
'   - For each RuleType, this function calls ApplyRule once using harmless dummy values.
'   - ApplyRule already contains a "Case Else" branch that returns:
'           "<field>: RuleType 'X' not implemented."
'     when a RuleType has no corresponding Case block.
'   - By probing ApplyRule directly, we verify that the VBA implementation exists for
'     every RuleType in configuration, independent of mapping rows or conditions.
'
' WHY THIS IS SAFE:
'   - The dummy values ("TestField", "TestValue") never touch real data.
'   - No movement, routing, or validation logic is triggered.
'   - This runs once per workflow, not per Unique Identifier (primary key) or mapping row.
'
' RESULT:
'   - If a RuleType is missing from the VBA engine, this function returns a configuration
'     error string. RunWorkflow will add it to fatalErrors and abort cleanly.
'   - If all RuleTypes are implemented, the function returns "".
'
'=========================================================================================
Public Function ValidateRuleEngine(ruleDict As Object) As String
    Dim ruleType As Variant
    Dim testErr As String
    Dim rt As String
    
    For Each ruleType In ruleDict.Keys
        
        ' Only probe RuleTypes marked Defined? = TRUE
        If ruleDict(ruleType) = True Then
        
            rt = CStr(ruleType)
            
            ' Probe the rule engine with dummy values
            testErr = ApplyRule(rt, "TestField", "TestValue")
            
            If InStr(1, testErr, "not implemented", vbTextCompare) > 0 Then
                ValidateRuleEngine = "RuleType '" & rt & "' is defined in the rule table but not implemented in the rule engine."
                Exit Function
            End If
        
        End If
    Next ruleType
    
    ValidateRuleEngine = ""
End Function


'=========================================================================================
' ValidateMovementTypeEngine
'
' PURPOSE:
'   Ensures that every MovementType defined in tblMovementTypes and marked Defined? = TRUE
'   has a corresponding implementation in the MovementType engine (ApplyMovementType).
'
'   This prevents a silent failure scenario where:
'       - A MovementType exists in the configuration table
'       - It is marked Defined? = TRUE
'       - But the corresponding Case block is missing in ApplyMovementType
'
' HOW IT WORKS:
'   - movementTypeDict contains every MovementType from tblMovementTypes with its Defined? flag.
'   - For each MovementType where Defined? = TRUE, this function calls ApplyMovementType once
'     using harmless dummy values.
'   - ApplyMovementType returns a special marker:
'           movementResult = "ENGINE_NOT_IMPLEMENTED"
'     when a MovementType has no corresponding Case block.
'
' WHY THIS IS SAFE:
'   - Dummy arguments (Nothing, "TEST") never touch real data.
'   - No movement, routing, or validation logic is triggered.
'   - This runs once per workflow, not per Unique Identifier (primary key) or mapping row.
'
' RESULT:
'   - If a MovementType is missing from the engine, this function returns a configuration
'     error string. RunWorkflow will add it to fatalErrors and abort cleanly.
'   - If all MovementTypes are implemented, the function returns "".
'
' LOCATION:
'   This belongs in modValidation alongside ValidateRuleEngine, as it validates the integrity
'   of the MovementType engine itself.
'=========================================================================================
Public Function ValidateMovementTypeEngine(movementTypeDict As Object) As String
    Dim mt As Variant
    Dim testResult As String
    Dim dummyRow As ListRow
    
    For Each mt In movementTypeDict.Keys
        
        ' Only probe MovementTypes that are marked Defined? = TRUE
        If movementTypeDict(mt) = True Then
            
            ' Probe the MovementType engine with dummy values.
            '   - tblSource, tblDest, srcDataRow are passed as Nothing.
            '   - pkvalue is a harmless test string.
            '   - isPrimaryKeyField is False (we don't care for the probe).
            testResult = ""
            Set dummyRow = ApplyMovementType(CStr(mt), Nothing, Nothing, Nothing, "TEST", "", False, testResult)
            
            ' If the engine reports ENGINE_NOT_IMPLEMENTED, the Case block is missing.
            If InStr(1, testResult, "ENGINE_NOT_IMPLEMENTED", vbTextCompare) > 0 Then
                ValidateMovementTypeEngine = "MovementType '" & CStr(mt) & "' is defined in tblMovementTypes but not implemented in ApplyMovementType."
                Exit Function
            End If
        End If
    Next mt
    
    ValidateMovementTypeEngine = ""
End Function

'=========================================================================================
' ValidateWriteModeEngine
'
' PURPOSE:
'   Ensures that every WriteMode defined in tblWriteModes and marked Defined? = TRUE
'   has a corresponding implementation in the WriteMode engine (ApplyWriteMode).
'
'   Prevents silent failures where:
'       - A WriteMode exists in configuration
'       - Defined? = TRUE
'       - But ApplyWriteMode has no Case block for it
'
' HOW IT WORKS:
'   - Calls ApplyWriteMode with destCell = Nothing (probe mode)
'   - ApplyWriteMode sets writeModeResult = "ENGINE_NOT_IMPLEMENTED"
'     when a Case block is missing
'
' SAFE:
'   - No real cells touched
'   - No movement logic triggered
'   - Runs once per workflow
'=========================================================================================
Public Function ValidateWriteModeEngine(writeModeDict As Object) As String
    Dim wm As Variant
    Dim testResult As String
    Dim dummyCell As Range
    
    For Each wm In writeModeDict.Keys
        
        If writeModeDict(wm) = True Then
            
            testResult = ""
            Call ApplyWriteMode(CStr(wm), "TEST_VALUE", "TEST_DEFAULT", dummyCell, testResult)

            
            If InStr(1, testResult, "ENGINE_NOT_IMPLEMENTED", vbTextCompare) > 0 Then
                ValidateWriteModeEngine = "WriteMode '" & CStr(wm) & "' is defined in tblWriteModes but not implemented in ApplyWriteMode."
                Exit Function
            End If
        End If
    Next wm
    
    ValidateWriteModeEngine = ""
End Function


' ============================================================
' ValidateConditionOperatorEngine
'
' PURPOSE:
'   Ensures that every operator marked Defined? = TRUE in
'   tblConditionalOperators is actually implemented in ConditionIsMet.
'
' HOW IT WORKS:
'   - Loops the operator dictionary (table-driven)
'   - Calls ConditionIsMet in probe mode
'   - Detects "ENGINE_NOT_IMPLEMENTED"
'
' ============================================================
Public Function ValidateConditionOperatorEngine(operatorDict As Object) As String
    Dim op As Variant
    Dim result As String

    For Each op In operatorDict.Keys
        If operatorDict(op) = True Then

            result = ""
            ' Probe mode: tblSource = Nothing
            Call ConditionIsMet(Nothing, Nothing, Nothing, CStr(op), result)

            If result = "ENGINE_NOT_IMPLEMENTED" Then
                ValidateConditionOperatorEngine = _
                    "Operator '" & CStr(op) & "' is defined in tblConditionalOperators but not implemented in ConditionIsMet."
                Exit Function
            End If

        End If
    Next op

    ValidateConditionOperatorEngine = ""
End Function



' ============================================================
' BuildMovementTypeDictionary
' ============================================================

Public Function BuildMovementTypeDictionary(tblMovementTypes As ListObject) As Object
    Dim dict As Object
    Dim r As ListRow
    Dim movementType As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each r In tblMovementTypes.ListRows
        movementType = Trim(CStr(r.Range.Columns(1).Value)) ' MovementType
        If movementType <> "" Then
            dict(movementType) = CBool(r.Range.Columns(3).Value) ' Defined?
        End If
    Next r
    
    Set BuildMovementTypeDictionary = dict
End Function

' ============================================================
' BuildConditionalOperatorDictionary
' ============================================================

Public Function BuildConditionalOperatorDictionary(tblConditionalOperators As ListObject) As Object
    Dim dict As Object
    Dim row As ListRow
    Dim op As String
    Dim isDefined As Boolean

    Set dict = CreateObject("Scripting.Dictionary")

    For Each row In tblConditionalOperators.ListRows
        op = Trim(CStr(row.Range.Columns(1).Value))          ' Operator
        isDefined = CBool(row.Range.Columns(3).Value)        ' Defined? column

        If op <> "" Then
            dict(op) = isDefined
        End If
    Next row

    Set BuildConditionalOperatorDictionary = dict
End Function

' ============================================================
' BuildWriteModeDictionary
' ============================================================

Public Function BuildWriteModeDictionary(tblWriteModes As ListObject) As Object
    Dim dict As Object
    Dim row As ListRow
    Dim mode As String
    Dim isDefined As Boolean

    Set dict = CreateObject("Scripting.Dictionary")

    For Each row In tblWriteModes.ListRows
        mode = Trim(CStr(row.Range.Columns(1).Value))
        isDefined = CBool(row.Range.Columns(3).Value)

        If mode <> "" Then
            dict(mode) = isDefined
        End If
    Next row

    Set BuildWriteModeDictionary = dict
End Function



' ============================================================
' BuildTableConfigurationDictionaries
'
'   Column 1 = TableName
'   Column 2 = PrimaryKeyField
'   Column 3 = EligibleForWorkflow (Boolean)
'   Column 4 = TableExecutionOrder (optional, blank = run last)
'
'   OUTPUT:
'       tablePKDict(tableName) = PrimaryKeyField          (original case)
'       tableEligibleDict(tableName) = Eligible? Boolean  (original case)
'       tableExecOrderDict(lcase(tableName)) = ExecOrder  (lowercase key)
'
'   WHY:
'       - groupedMappings.Keys are lowercase
'       - therefore tableExecOrderDict must also be keyed lowercase
'       - PK + eligibility must remain original case for validation
' ============================================================
Public Sub BuildTableConfigurationDictionaries( _
        tblWorkflowTables As ListObject, _
        ByRef tablePKDict As Object, _
        ByRef tableEligibleDict As Object, _
        ByRef tableExecOrderDict As Object)

    Dim r As ListRow
    Dim tableName As String
    Dim pkField As String
    Dim eligible As Boolean
    Dim execOrderVal As Variant
    Dim execOrder As Long

    Set tablePKDict = CreateObject("Scripting.Dictionary")
    tablePKDict.CompareMode = vbTextCompare

    Set tableEligibleDict = CreateObject("Scripting.Dictionary")
    tableEligibleDict.CompareMode = vbTextCompare

    Set tableExecOrderDict = CreateObject("Scripting.Dictionary")
    tableExecOrderDict.CompareMode = vbTextCompare

    For Each r In tblWorkflowTables.ListRows
        tableName = Trim(CStr(r.Range.Columns(1).Value))
        pkField = Trim(CStr(r.Range.Columns(2).Value))
        eligible = CBool(r.Range.Columns(3).Value)
        execOrderVal = r.Range.Columns(4).Value

        If tableName <> "" Then
            ' original-case dictionaries (used elsewhere)
            tablePKDict(tableName) = pkField
            tableEligibleDict(tableName) = eligible

            ' robust normalisation for execution order
            execOrder = 999999   ' default = run last
            If Not IsError(execOrderVal) Then
                If Not IsEmpty(execOrderVal) Then
                    If Len(Trim(CStr(execOrderVal))) > 0 Then
                        If IsNumeric(execOrderVal) Then
                            execOrder = CLng(execOrderVal)
                        End If
                    End If
                End If
            End If

            ' store under lowercase key to match groupedMappings
            tableExecOrderDict(LCase(tableName)) = execOrder
        End If
    Next r
End Sub

