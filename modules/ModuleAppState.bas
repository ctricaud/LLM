Attribute VB_Name = "ModuleAppState"
'=== modAppState ===
Option Explicit

Private Type TAppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calc As XlCalculation
    DisplayStatusBar As Boolean
End Type

Private AppState As TAppState

Public Sub BeginAppState(Optional ByVal showStatus As Boolean = True)
    
    With Application
        AppState.ScreenUpdating = .ScreenUpdating
        AppState.EnableEvents = .EnableEvents
        AppState.Calc = .Calculation
        AppState.DisplayStatusBar = .DisplayStatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        '.DisplayStatusBar = showStatus
        If .Calculation <> xlCalculationManual Then .Calculation = xlCalculationManual
        '.StatusBar = "Traitement en cours…"
    End With
End Sub

Public Sub EndAppState()
    
    On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        '.DisplayStatusBar = AppState.DisplayStatusBar
        '.StatusBar = False
    End With
End Sub

' Patron d’utilisation
Public Sub UsingAppState(ByVal actionName As String, ByVal proc As String)
    On Error GoTo EH
    BeginAppState
    Application.StatusBar = "Exécution: " & actionName
    ' >>> appelle la procédure cible par Application.Run pour éviter les références croisées
    Application.Run proc
    GoTo Finally

EH:
    LogError proc, Erl, Err.Number, Err.Description
Finally:
    EndAppState
End Sub

