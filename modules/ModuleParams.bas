Attribute VB_Name = "ModuleParams"
'=== modParams ===
Option Explicit

Public Function GetParam(ByVal key As String, Optional ByVal defaultValue As Variant) As Variant
    Dim v As Variant
    v = GetParamFromNames(key)
    If IsError(v) Or v = vbNullString Then v = GetParamFromSheet(key)
    If IsError(v) Or v = vbNullString Then
        If Not IsMissing(defaultValue) Then
            GetParam = defaultValue
        Else
            GetParam = vbNullString
        End If
    Else
        GetParam = v
    End If
End Function

Private Function GetParamFromNames(ByVal key As String) As Variant
    On Error GoTo EH
    Dim nm As Name, mapKey As String
    mapKey = KeyToName(key)
    For Each nm In ThisWorkbook.Names
        If StrComp(nm.Name, mapKey, vbTextCompare) = 0 Then
            GetParamFromNames = nm.RefersToRange.value
            Exit Function
        End If
    Next
    GetParamFromNames = CVErr(xlErrNA)
    Exit Function
EH:
    GetParamFromNames = CVErr(xlErrNA)
End Function

Private Function GetParamFromSheet(ByVal key As String) As Variant
    On Error GoTo EH
    Dim ws As Worksheet, f As Range
    Set ws = ThisWorkbook.Worksheets(SHEET_PARAMS)
    Set f = ws.Range("A:A").Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        GetParamFromSheet = f.Offset(0, 1).value
    Else
        GetParamFromSheet = CVErr(xlErrNA)
    End If
    Exit Function
EH:
    GetParamFromSheet = CVErr(xlErrNA)
End Function

Private Function KeyToName(ByVal key As String) As String
    Select Case key
        Case PARAM_EXPORT_DIR: KeyToName = NAME_EXPORT_DIR
        Case PARAM_CURRENT_YEAR: KeyToName = NAME_CURRENT_YEAR
        Case PARAM_LODGINGS: KeyToName = NAME_LODGINGS
        Case Else: KeyToName = key
    End Select
End Function

