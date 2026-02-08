VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub changeGraph()
    '----------------------------------------------------------------------
    'Procédure appelé lorsque les combo box sont modifiées
    '----------------------------------------------------------------------
    Dim NomTableau, nomSerie As String
    NomTableau = "statsAirbnb" + cbLogement.Text
    nomSerie = cbSerie.Text
    
    Dim ch As Chart
    On Error Resume Next
    Set ch = ActiveSheet.ChartObjects("MonGraphique").Chart
    If Err Then Exit Sub
    
    ' Sélectionner la première série
    Dim s As Series
    Set s = ch.SeriesCollection(1)
    
    '----------------------------------------------------------------------
    'On met à jour les données du graphique
    '----------------------------------------------------------------------
    s.Name = nomSerie + Chr(10) + cbLogement.Text
    
    s.values = "=" + NomTableau + "[" + nomSerie + "]"
    s.XValues = "=" + NomTableau + "[Date]"
    
    With ActiveSheet.ChartObjects("MonGraphique").Chart
    .Axes(xlCategory).CategoryType = xlTimeScale
    With .Axes(xlCategory)
        .CategoryType = xlTimeScale  ' Pour forcer l'axe en mode date
        .TickLabels.NumberFormat = "dd/mm/yyyy"  ' Format d'affichage
    End With
End With
End Sub




Private Sub cbLogement_Change()
    changeGraph
End Sub


Private Sub cbSerie_Change()
    changeGraph
End Sub


Private Sub Worksheet_Activate()
 
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


