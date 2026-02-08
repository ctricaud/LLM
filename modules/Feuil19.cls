VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 6 And Target.value = "X" Then
        'On supprime la ligne
        Range("Notifs").ListObject.ListRows(Target.row - 2).Delete
        Cancel = True
    End If
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


