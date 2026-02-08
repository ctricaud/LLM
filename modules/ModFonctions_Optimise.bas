Attribute VB_Name = "ModFonctions_Optimise"
Option Explicit

' ========= Optimisations d'indexation =========
' Construction d'un dictionnaire [valeur -> index] à partir d'un Tableau Structuré nommé
' Utilisé pour remplacer des recherches linéaires répétées (indexRange) par des recherches O(1).
'
' Exemple :
'   Dim dLog As Object: Set dLog = BuildIndexCache("Logements", 1)
'   iLog = IndexOf(dLog, "Apollinaire")
'
' NB : On reconstruit le cache à chaque appel volontaire (léger coût fixe),
'      mais on évite des milliers d'appels à indexRange() dans les boucles.
' =============================================

Public Function BuildIndexCache(ByVal nomPlage As String, ByVal col As Long) As Object
    'Retourne un Scripting.Dictionary: clé = valeur de la colonne "col", valeur = index (1-based dans le TS)
    Dim lo As ListObject
    Dim arr As Variant
    Dim i As Long
    Dim d As Object
    
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    On Error GoTo EH
    Set lo = Range(nomPlage).ListObject
    arr = lo.DataBodyRange.Value
    
    For i = 1 To UBound(arr, 1)
        'clé au format texte pour stabilité
        d(CStr(arr(i, col))) = i
    Next i
    
    Set BuildIndexCache = d
    Exit Function
EH:
    ' En cas de TS vide (ou création en cours), renvoie un dict vide
    Set BuildIndexCache = d
End Function

Public Function IndexOf(ByVal d As Object, ByVal key As Variant) As Long
    'Retourne l'index si trouvé, 0 sinon
    If d Is Nothing Then
        IndexOf = 0
    ElseIf d.Exists(CStr(key)) Then
        IndexOf = CLng(d(CStr(key)))
    Else
        IndexOf = 0
    End If
End Function

' ====== Version originale conservée pour compatibilité ======
' Recherche linéaire (O(n)) – laissée telle quelle pour ne pas casser l'existant.
Public Function indexRange(Nom, colonne, valeur) As Integer
   'Retourne l'index d'une valeur dans un tableau
   Dim i As Integer
   Dim R As Range
   
    Set R = Range(Nom)
    indexRange = 0      'Valeur si pas trouvé
    
    For i = 1 To R.ListObject.ListRows.Count
        If R(i, colonne).Value = valeur Then
            indexRange = i
            Exit For
        End If
    Next i
End Function
