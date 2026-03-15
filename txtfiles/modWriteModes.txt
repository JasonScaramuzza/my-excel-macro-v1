Attribute VB_Name = "modWriteModes"

' ================================================================================
'  © 2026 [Scaramuzza]. All rights reserved.
'  WorkflowXL – Version 1.0 (14 March 2026)
'
'  Module: [modWriteModes]
'  Purpose: [Contains the WriteModeEngine]
' ================================================================================


' ============================================================
' APPLY WRITEMODE (WriteMode Engine)
'
' PURPOSE:
'   Decides how to write a single cell based on WriteMode.
'
'   This also supports an engine-probe mode via the
'   optional writeModeResult parameter, allowing the engine
'   to detect unimplemented WriteModes safely.
'
'   - Normal calls (from PerformMovementForRow):
'         • destCell is a real Range
'         • writeModeResult is ignored by caller
'
'   - Probe calls (from ValidateWriteModeEngine):
'         • destCell is passed as Nothing
'         • No cell is touched
'         • writeModeResult is set to:
'               ""                       if implemented
'               "ENGINE_NOT_IMPLEMENTED" if missing
' ============================================================
Public Sub ApplyWriteMode(ByVal writeMode As String, _
                          ByVal sourceValue As Variant, _
                          ByVal defaultValue As Variant, _
                          ByRef destCell As Range, _
                          Optional ByRef writeModeResult As String)

    Dim currentVal As Variant
    Dim existing As String

    writeModeResult = ""   ' Default: assume implemented

    ' Probe vs real mode:
    '   - Real:  destCell Is NOT Nothing - we can read/write the cell.
    '   - Probe: destCell Is Nothing     - we must NOT touch any cell.
    If Not destCell Is Nothing Then
        currentVal = destCell.Value
    End If

    Select Case writeMode

        Case "", "Overwrite"
            If Not destCell Is Nothing Then
                destCell.Value = sourceValue
            End If

        Case "FallbackDefaultValueIfBlank"
            If Not destCell Is Nothing Then
                If Trim(CStr(sourceValue)) = "" Then
                    destCell.Value = defaultValue
                Else
                    destCell.Value = sourceValue
                End If
            End If

        Case "Timestamp"
            If Not destCell Is Nothing Then
                destCell.Value = Now
            End If

        Case "Ignore"
            ' Do nothing in both real and probe mode

        Case "OnlyIfBlank"
            If Not destCell Is Nothing Then
                If Trim(CStr(currentVal)) = "" Then
                    destCell.Value = sourceValue
                End If
            End If

        Case "AppendData"
            ' Append non-blank sourceValue to existing content, on a new line.
            existing = CStr(currentVal)

            If Not destCell Is Nothing Then
                If Trim(CStr(sourceValue)) <> "" Then
                    If Len(existing) = 0 Then
                        destCell.Value = sourceValue
                    Else
                        destCell.Value = existing & vbLf & sourceValue
                    End If
                End If
            End If

        Case Else
            ' Unknown WriteMode:
            '   - Real mode: safest is to do nothing.
            '   - Probe mode: self-report as not implemented.
            If destCell Is Nothing Then
                writeModeResult = "ENGINE_NOT_IMPLEMENTED"
            End If
    End Select
End Sub
