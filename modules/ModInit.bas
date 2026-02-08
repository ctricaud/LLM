Attribute VB_Name = "ModInit"
'Ce module permet d'initialiser des variables
'Au lancement du classeur
Option Explicit

Public idxResas As Object

Public dicLog As Object     'La liste des logements actifs
Public dicLogA As Object    'Leur position dans la liste
Public dicLogB As Object    'La liste des ligements dans le budget
Public dicLogs As Object    'Tous les logements
Public ListeLogements       'La liste de tous les logements

Public dicSrc As Object
Public ListeSources         'La liste de toutes les sources de réservattions
Sub CalculIdxResas()
'---------------------------------------------------------
'Mise à jour des cononnes ed listeRésas
'---------------------------------------------------------
    Dim i As Integer
    
    Set idxResas = New Dictionary
    For i = 1 To Range("ListeRésas").ListObject.ListColumns.Count
        idxResas(Range("ListeRésas").ListObject.ListColumns(i).Name) = i
    Next i
    
End Sub

Sub Init(Optional Force = False)

    If TypeName(dicLog) = "Dictionary" And Not Force Then Exit Sub
    '-------------------------------------------------------------------
    'Le dictionnaire des colonnes du tableau résservation
    '-------------------------------------------------------------------
    
    CalculIdxResas
    
    '-------------------------------------------------------------------
    'Les dictionnaires logements
    '-------------------------------------------------------------------
    Dim i As Integer
    
    ListeLogements = Feuil5.Range("Logements")
    
    Set dicLog = New Dictionary
    Set dicLogA = New Dictionary
    Set dicLogB = New Dictionary
    Set dicLogs = New Dictionary
    
    Dim n, p
    For i = 1 To UBound(ListeLogements)
        dicLogs(ListeLogements(i, 1)) = i
        
        If ListeLogements(i, 8) Then
            n = n + 1
            dicLog(ListeLogements(i, 1)) = n
            dicLogA(ListeLogements(i, 1)) = i
        End If
        
        If ListeLogements(i, 2) <> "" Then
            p = p + 1
            dicLogB(ListeLogements(i, 1)) = n
        End If
    Next i
    
    '-------------------------------------------------------------------
    'Les dictionnaires des sources
    '-------------------------------------------------------------------
    ListeSources = Feuil5.Range("Sources")
    
    Set dicSrc = New Dictionary
    
    For i = 1 To UBound(ListeSources)
        dicSrc(ListeSources(i, 1)) = i
    Next i
    
End Sub



