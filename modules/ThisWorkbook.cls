VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Declare PtrSafe Function w_commandline Lib "kernel32.dll" Alias "GetCommandLineW" () As LongPtr
Private Declare PtrSafe Function w_strlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub w_memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal size As LongPtr)

Public Function GetCommandLine() As String
    GetCommandLine = String$(w_strlen(w_commandline()), 0)
    w_memcpy ByVal StrPtr(GetCommandLine), ByVal w_commandline(), LenB(GetCommandLine)
End Function

Private Sub Workbook_Activate()
    Init
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   Sauvegarde
End Sub

Private Sub Workbook_Open()
    'RAZTableau "Planificateur"
   'ApplicationOntime Date - (Hour(Now) >= 7) + TimeValue("07:00:00"), "PlanificationGlobale"
   'ApplicationOntime Now + TimeValue("00:02:00"), "NotificationsPulse"
   'ApplicationOntime Date + 1 + ((8 - (Date Mod 7)) Mod 7) + TimeSerial(7, 30, 0), "alerte2J"
   
   MAJTravaux
    
    Dim monparam As Variant                      'déclare une variable
    Dim i As Integer
    
   
    monparam = GetCommandLine                    'affecte la valeur de la ligne de commande
     
    If Not IsNull(monparam) Then
        If Len(monparam) > 0 Then
           
            i = InStr(monparam, "/cmd/")
            If i > 0 Then
                monparam = Mid(monparam, i + 5)
                Application.Run monparam
            End If
        End If
    End If
End Sub

