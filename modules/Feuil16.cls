VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_Change()
    'On met à jour le logement
    Range("LogementChannel") = Feuil16.ComboBox1.Text
    majCalculChannel
End Sub


Private Sub Worksheet_Activate()
    With Feuil16.ComboBox1
        .Clear
        .List = Feuil5.Range("Logements[Logements]").value
        .AddItem "Maury + Apollinaire"
        .AddItem "Tous les logements"
        .ListIndex = 5
    End With
   majCalculChannel
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

