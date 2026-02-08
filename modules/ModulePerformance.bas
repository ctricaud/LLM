Attribute VB_Name = "ModulePerformance"
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Sub TestChrono()
    Dim t0 As Long
    t0 = ChronoStart()
    
    ' Simuler une tâche
    Dim i As Long
    For i = 1 To 1000
        DoEvents
    Next i
    
    Call ChronoStop(t0, "TestChrono")
End Sub

'--- Lance le chrono et renvoie l'heure en ms
Public Function ChronoStart() As Long
    ChronoStart = GetTickCount()
End Function

Public Sub ChronoStop(ByVal t0 As Long, ByVal procName As String)
    Dim duree As Long
    duree = GetTickCount() - t0
    
    If duree > PERF_THRESHOLD Then
        WriteLog procName, duree
    End If
End Sub
Private Sub WriteLog(ByVal procName As String, ByVal duree As Long)
    Dim fso As Object, ts As Object
    Dim logFile As String
    
    logFile = CheminFichier & "\PerfLog.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logFile, 8, True) ' 8 = Append, True = créer si inexistant
    
    ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & procName & " | " & duree & " ms"
    ts.Close
End Sub
