Attribute VB_Name = "ModCalcul_Optimise"
Option Explicit

' ===================================================================
' CalculJourV2 — Version optimisée en temps d'exécution
'
' Principes d'optimisation :
' 1) Suppression des recherches répétées via indexRange() (O(n))
'    -> Remplacées par des dictionnaires (O(1)) construits une seule fois.
' 2) Zéro accès cellule/Range dans les boucles : tout est en tableaux Variants.
' 3) Typage systématique en Long (évite conversions et dépassements Integer).
' 4) Calculs précomposés (prix/nuit, bornes de dates) hors des boucles internes.
'
' Signature : identique au CalculJour existant (renvoie un Variant 3D)
' Dimensions retournées : (jour de DateDébut->DateFin, nbLogements*2, nbSources)
'   - 1..nbLogements  : prix net reçu (par nuit)
'   - nbLogements+1.. : prix payé par le client (par nuit)
' ===================================================================
Public Function CalculJourV2(Optional ByVal anDebut As Long = 2023, Optional ByVal anFin As Long = 2030) As Variant
    Dim dateDebut As Date, dateFin As Date
    Dim loRes As ListObject
    Dim arrRes As Variant
    Dim arrLo As Variant, arrSrc As Variant
    Dim nbLogements As Long, nbSources As Long
    Dim dLog As Object, dSrc As Object
    Dim iRes As Long, j As Long
    Dim iLog As Long, iSrc As Long
    Dim jStart As Long, jEnd As Long
    Dim prixNuit As Currency, prixNuitClient As Currency
    Dim tabJour As Variant
    Dim loLog As ListObject, loSrc As ListObject
    Dim lbJour As Long, ubJour As Long
    
    If Not IsNumeric(anDebut) Or Not IsNumeric(anFin) Then
        MsgBox "Erreur dans la saisie des dates", vbExclamation
        Exit Function
    End If
    
    If anFin < anDebut Then
        Dim tmp As Long
        tmp = anDebut: anDebut = anFin: anFin = tmp
    End If
    
    dateDebut = DateSerial(anDebut, 1, 1)
    dateFin = DateSerial(anFin, 12, 31)
    
    ' --- Structures de référence (Logements, Sources) ---
    Set loLog = Range("Logements").ListObject
    Set loSrc = Range("Sources").ListObject
    nbLogements = loLog.ListRows.Count
    nbSources = loSrc.ListRows.Count
    
    ' Dictionnaires de lookup (clé = texte de la colonne 1)
    Set dLog = ModFonctions_Optimise.BuildIndexCache("Logements", 1)
    Set dSrc = ModFonctions_Optimise.BuildIndexCache("Sources", 1)
    
    ' --- Réservations ---
    Set loRes = Range("ListeRésas").ListObject
    If loRes.ListRows.Count = 0 Then
        ' Tableau 3D vide mais correctement dimensionné
        ReDim tabJour(CLng(dateDebut) To CLng(dateFin), 1 To nbLogements * 2, 1 To nbSources)
        CalculJourV2 = tabJour
        Exit Function
    End If
    arrRes = loRes.DataBodyRange.Value
    
    ' --- Tableau résultat : Variant 3D (dates en base 0 = num série) ---
    lbJour = CLng(dateDebut)
    ubJour = CLng(dateFin)
    ReDim tabJour(lbJour To ubJour, 1 To nbLogements * 2, 1 To nbSources)
    
    ' Colonnes attendues dans ListeRésas (adapter si besoin) :
    '  1 = Logement (texte)
    '  2 = Source (texte)
    '  3 = Date début (Date/num série)
    '  4 = Nb Nuits (Long)
    '  5 = PrixNuitClient (Currency) -> prix payé par le client
    ' 10 = MontantVersé (Currency)   -> total reçu (frais déduits) pour le séjour
    ' Si vos colonnes diffèrent, ajustez les index ci-dessous.
    Const COL_LOG As Long = 1
    Const COL_SRC As Long = 2
    Const COL_DATE As Long = 3
    Const COL_NBNUITS As Long = 4
    Const COL_PRIXNUITCLIENT As Long = 5
    Const COL_MONTANT_VERSE As Long = 10
    
    Dim nbRes As Long
    nbRes = UBound(arrRes, 1)
    
    ' Boucle réservations -> répartition par jour
    For iRes = 1 To nbRes
        iLog = ModFonctions_Optimise.IndexOf(dLog, arrRes(iRes, COL_LOG))
        If iLog = 0 Then GoTo NextRes
        
        iSrc = ModFonctions_Optimise.IndexOf(dSrc, arrRes(iRes, COL_SRC))
        If iSrc = 0 Then GoTo NextRes
        
        ' Borne de dates de la résa
        jStart = CLng(arrRes[iRes, COL_DATE])
        jEnd = jStart + CLng(arrRes[iRes, COL_NBNUITS]) - 1
        
        ' Clamping dans l'intervalle demandé
        If jEnd < lbJour Or jStart > ubJour Then GoTo NextRes
        If jStart < lbJour Then jStart = lbJour
        If jEnd > ubJour Then jEnd = ubJour
        
        ' Prix/nuit pré-calculés
        If CLng(arrRes(iRes, COL_NBNUITS)) > 0 Then
            prixNuit = CCur(arrRes(iRes, COL_MONTANT_VERSE)) / CLng(arrRes(iRes, COL_NBNUITS))
        Else
            prixNuit = 0
        End If
        prixNuitClient = CCur(arrRes(iRes, COL_PRIXNUITCLIENT))
        
        ' Remplissage pour chaque jour de la réservation
        For j = jStart To jEnd
            tabJour(j, iLog, iSrc) = CCur(tabJour(j, iLog, iSrc)) + prixNuit
            tabJour(j, iLog + nbLogements, iSrc) = CCur(tabJour(j, iLog + nbLogements, iSrc)) + prixNuitClient
        Next j
NextRes:
    Next iRes
    
    CalculJourV2 = tabJour
End Function

' ====== Wrapper de compatibilité ======
' Vous pouvez remplacer les appels existants à CalculJour par CalculJourV2
' sans changer le reste du code.
Public Function CalculJour(Optional ByVal anDebut As Long = 2023, Optional ByVal anFin As Long = 2030) As Variant
    CalculJour = CalculJourV2(anDebut, anFin)
End Function
