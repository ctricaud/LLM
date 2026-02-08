VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public MoisFiltre As Integer, AnneeFiltre As Integer

Sub CheckFiltre()
    If FiltreMois.Caption <> "Mois en cours" And MoisFiltre = 0 Then
        MoisFiltre = CInt(Left(FiltreMois.Caption, 2))
        AnneeFiltre = CInt(Mid(FiltreMois.Caption, 4))
    End If
End Sub

Sub Filtrage()
    If MoisFiltre <> 0 Then
        Dim dStart As Long
        Dim dEnd As Long
        
        dStart = DateSerial(AnneeFiltre, MoisFiltre, 1)
        dEnd = DateSerial(AnneeFiltre, MoisFiltre + 1, 0)
        
        Range("ListeRésas").AutoFilter Field:=3, _
                                       Criteria1:=">=" & CStr(dStart), Operator:=xlAnd, Criteria2:="<=" & CStr(dEnd)
    End If
End Sub



Private Sub FiltreMois_Click()
    If mois = 0 Then
        MoisFiltre = Month(Date)
        AnneeFiltre = Year(Date)
        
        FiltreMois.Caption = Format(MoisFiltre, "0#") + "/" + CStr(AnneeFiltre)
        Filtrage
    End If
End Sub

Public Sub RAZFiltres_Click()
    '+++++++++++++++++++++++++++++++++++++++++++++++
    '+ On efface tous les filtres
    '+++++++++++++++++++++++++++++++++++++++++++++++
    With Range("ListeRésas").ListObject
        If Not .AutoFilter Is Nothing _
           Then .AutoFilter.ShowAllData
    End With
    
 
    
    'On met à jour le texte du bouton de sélection du mois
    FiltreMois.Caption = "Mois en cours"
    MoisFiltre = 0
End Sub

Private Sub SpinButton1_Change()

End Sub

Private Sub SpinButton1_SpinDown()
    CheckFiltre
    If MoisFiltre <> 0 Then
        MoisFiltre = MoisFiltre - 1
        If MoisFiltre = 0 Then
            MoisFiltre = 12
            AnneeFiltre = AnneeFiltre - 1
        End If
        
        FiltreMois.Caption = Format(MoisFiltre, "0#") + "/" + CStr(AnneeFiltre)
    End If
    
    Filtrage

End Sub

Private Sub SpinButton1_SpinUp()
    CheckFiltre
    If MoisFiltre <> 0 Then
        MoisFiltre = MoisFiltre + 1
        If MoisFiltre = 13 Then
            MoisFiltre = 1
            AnneeFiltre = AnneeFiltre + 1
        End If
        
        FiltreMois.Caption = Format(MoisFiltre, "0#") + "/" + CStr(AnneeFiltre)
    End If
    
    Filtrage
End Sub

Private Sub Worksheet_Activate()
    'majStatut
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub


