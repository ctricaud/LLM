Attribute VB_Name = "ModuleError"
'=== modErrLog ===
Option Explicit

Public Sub EnsureLogSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:D1").value = Array("Timestamp", "Proc", "ErrNum", "Message")
    End If
End Sub

Public Sub LogError(ByVal proc As String, ByVal lineNum As Long, ByVal errNum As Long, ByVal msg As String)
    On Error Resume Next
    Call EnsureLogSheet
    Dim ws As Worksheet, nextRow As Long
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = proc & IIf(lineNum > 0, " [Erl=" & lineNum & "]", "")
    ws.Cells(nextRow, 3).value = errNum
    ws.Cells(nextRow, 4).value = msg
    Debug.Print "ERR " & proc & " #" & errNum & " @L" & lineNum & " : " & msg
End Sub

Public Sub LogInfo(ByVal msg As String)
    On Error Resume Next
    Call EnsureLogSheet
    Dim ws As Worksheet, nextRow As Long
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = "INFO"
    ws.Cells(nextRow, 4).value = msg
    Debug.Print "INFO " & msg
End Sub

