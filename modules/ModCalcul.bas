Attribute VB_Name = "ModCalcul"

Option Explicit

Sub CalculAnnulations()
    '1 On prépare les tableaux
    Dim T As Variant
    Dim U As Variant
    
    T = Range("ListeRésas")
    U = Range("Annulation")
    
    Dim Nuits(1 To 13, 2025 To 2030)
    Dim NuitsBooking(1 To 13, 2025 To 2030)
    Dim NuitsAirbnb(1 To 13, 2025 To 2030)
   Dim Annulations(1 To 13, 2025 To 2030)
    Dim AnnulationsBooking(1 To 13, 2025 To 2030)
    Dim AnnulationsAirbnb(1 To 13, 2025 To 2030)
    
    Dim i, j As Long
        
    '2 On calcule les nuits à partir de la date de réservation
    For i = 1 To UBound(T)
        If Year(T(i, 17)) > 2024 Then
            Nuits(Month(T(i, 17)), Year(T(i, 17))) = Nuits(Month(T(i, 17)), Year(T(i, 17))) + T(i, 4)
            Select Case T(i, 2)
                Case "Airbnb": NuitsAirbnb(Month(T(i, 17)), Year(T(i, 17))) = NuitsAirbnb(Month(T(i, 17)), Year(T(i, 17))) + T(i, 4)
                Case "Booking": NuitsBooking(Month(T(i, 17)), Year(T(i, 17))) = NuitsBooking(Month(T(i, 17)), Year(T(i, 17))) + T(i, 4)
            End Select
        End If
    Next i
    
    '3 On calcule les nuits annulées à partir de la date de réservation
    For i = 1 To UBound(U)
        If Year(U(i, 8)) > 2024 Then
            Annulations(Month(U(i, 8)), Year(U(i, 8))) = Annulations(Month(U(i, 8)), Year(U(i, 8))) + U(i, 4)
            Select Case U(i, 2)
                Case "Airbnb": AnnulationsAirbnb(Month(U(i, 8)), Year(U(i, 8))) = AnnulationsAirbnb(Month(U(i, 8)), Year(U(i, 8))) + U(i, 4)
                Case "Booking": AnnulationsBooking(Month(U(i, 8)), Year(U(i, 8))) = AnnulationsBooking(Month(U(i, 8)), Year(U(i, 8))) + U(i, 4)
            End Select
        End If
    Next i
    
    '4 On fait les totaux
    For i = 1 To 12
        For j = 2025 To 2030
            Nuits(13, j) = Nuits(13, j) + Nuits(i, j)
            NuitsAirbnb(13, j) = NuitsAirbnb(13, j) + NuitsAirbnb(i, j)
            NuitsBooking(13, j) = NuitsBooking(13, j) + NuitsBooking(i, j)
            Annulations(13, j) = Annulations(13, j) + Annulations(i, j)
            AnnulationsAirbnb(13, j) = AnnulationsAirbnb(13, j) + AnnulationsAirbnb(i, j)
            AnnulationsBooking(13, j) = AnnulationsBooking(13, j) + AnnulationsBooking(i, j)
        Next j
    
    Next i
    
    '5 On met à jour le tableau avec les valeurs obtenues
    Dim R As Variant
    Dim texte As String
    R = Range("TauxAnnulation")
    
    For i = 1 To 13
        For j = 2025 To 2030
            texte = ""
            '4.1 On calcule pour le total
            If Annulations(i, j) <> 0 Then
                texte = CStr(CInt(Annulations(i, j) / (Nuits(i, j) + Annulations(i, j)) * 100)) + "% ("
                    If AnnulationsAirbnb(i, j) <> 0 Then
                        texte = texte + CStr(CInt(AnnulationsAirbnb(i, j) / (NuitsAirbnb(i, j) + AnnulationsAirbnb(i, j)) * 100)) + " , "
                    Else
                        texte = texte + "- , "
                    End If
                    If AnnulationsBooking(i, j) <> 0 Then
                        texte = texte + CStr(CInt(AnnulationsBooking(i, j) / (NuitsBooking(i, j) + AnnulationsBooking(i, j)) * 100))
                    Else
                        texte = texte + "-"
                    End If
                texte = texte + ")"
            End If
            R(i, j - 2024) = texte
        Next j
    Next i
    
    '5 On met à jour le tableau
    Range("TauxAnnulation") = R
    
End Sub


' -------------------------------------------------------------------------
' Procédure : CalculProjections
' Auteur    : Gemini
' Date      : 23/11/2025 (Mise à jour : Ajout ligne N-2)
' Objectif  : Remplit le tableau "Projections" (M+1 à M+5) pour N, N-1 et N-2.
'             FILTRE : Uniquement les logements "Maury" et "Apollinaire".
' -------------------------------------------------------------------------
Sub CalculProjections()
    ' -- Déclaration des variables --
    Dim wsResa As Worksheet, wsAccueil As Worksheet
    Dim loResas As ListObject, loProj As ListObject
    Dim vDataResas As Variant
    Dim i As Long, j As Long, k As Long
    Dim dateRef As Date, dateDebutSejour As Date, dateBooking As Date
    Dim dMoisCible As Date
    Dim montantTotal As Double, montantLigne As Double
    Dim sLocation As String
    Dim sLibelleLigne As String
    
    ' Index des colonnes
    Dim colBooking As Long, colDebutSejour As Long, colMontant As Long, colLocation As Long
    
    ' -- CONSTANTES DES EN-TÊTES (A vérifier dans votre fichier) --
    Const NOM_COL_LOCATION As String = "Location"       ' La colonne avec "Maury", "Apollinaire"
    Const NOM_COL_BOOKING As String = "booking_date"
    Const NOM_COL_DEBUT As String = "Date Début"       ' Vérifiez l'orthographe
    Const NOM_COL_MONTANT As String = "Versement"         ' La colonne prix
    Const NB_AN_CALCUL As Integer = 3
    
    ' -- Gestion des erreurs et Optimisation --
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. Initialisation des feuilles et tableaux
    Set wsResa = ThisWorkbook.Sheets("Réservations")
    Set wsAccueil = ThisWorkbook.Sheets("Accueil")
    
    On Error Resume Next
    Set loResas = wsResa.ListObjects("ListeRésas")
    Set loProj = wsAccueil.ListObjects("Projections")
    On Error GoTo ErrorHandler
    
    If loResas Is Nothing Or loProj Is Nothing Then
        MsgBox "Erreur : Impossible de trouver les tableaux 'ListeRésas' ou 'Projections'.", vbCritical
        GoTo FinProgramme
    End If
    
    ' 2. Chargement des données en mémoire
    If loResas.ListRows.Count = 0 Then GoTo FinProgramme
    vDataResas = loResas.DataBodyRange.value
    
    ' 3. Récupération des index de colonnes
    On Error Resume Next
    colLocation = loResas.ListColumns(NOM_COL_LOCATION).Index
    colBooking = loResas.ListColumns(NOM_COL_BOOKING).Index
    colDebutSejour = loResas.ListColumns(NOM_COL_DEBUT).Index
    colMontant = loResas.ListColumns(NOM_COL_MONTANT).Index
    On Error GoTo ErrorHandler
    
    If colLocation = 0 Or colBooking = 0 Or colDebutSejour = 0 Or colMontant = 0 Then
        MsgBox "Erreur colonnes introuvables. Vérifiez les constantes au début du code.", vbCritical
        GoTo FinProgramme
    End If

    ' 4. Structure du tableau de projection (Garantir 3 lignes maintenant)
    If loProj.ListRows.Count < 3 Then
        Do While loProj.ListRows.Count < 3
            loProj.ListRows.Add
        Loop
    End If
    
    ' 5. Boucle principale de calcul (1 à 3 pour N, N-1, N-2)
    For i = 1 To NB_AN_CALCUL
        
        ' Définition de la date pivot et du libellé selon la ligne
        ' i = 1 -> décalage 0 (Année N)
        ' i = 2 -> décalage -1 (Année N-1)
        ' i = 3 -> décalage -2 (Année N-2)
        
        dateRef = DateAdd("yyyy", -(i - 1), Date)
        
        If i = 1 Then
            sLibelleLigne = "Proj " & Format(dateRef, "dd/mm/yyyy")
        Else
            sLibelleLigne = "N" & (1 - i) & " " & Format(dateRef, "dd/mm/yyyy")
        End If
        
        ' Ecriture du libellé en colonne 1
        loProj.DataBodyRange(i, 1).value = sLibelleLigne
        
        ' Boucle sur les colonnes M+1 à M+5
        For j = 1 To 5
            
            ' Calcul du mois cible (Mois de dateRef + j)
            dMoisCible = DateSerial(Year(dateRef), Month(dateRef) + j, 1)
            montantTotal = 0
            
            ' -- Parcours des données --
            For k = LBound(vDataResas, 1) To UBound(vDataResas, 1)
                
                ' A. Récupération du nom du logement
                sLocation = UCase(Trim(CStr(vDataResas(k, colLocation))))
                
                ' B. FILTRE : Uniquement Maury ou Apollinaire
                If sLocation = "MAURY" Or sLocation = "APOLLINAIRE" Then
                    
                    ' Vérification dates valides
                    If IsDate(vDataResas(k, colBooking)) And IsDate(vDataResas(k, colDebutSejour)) Then
                        dateBooking = vDataResas(k, colBooking)
                        dateDebutSejour = vDataResas(k, colDebutSejour)
                        
                        ' C. Critère date de réservation (<= Date Ref)
                        If dateBooking <= dateRef Then
                            
                            ' D. Critère date de séjour (Mois cible)
                            If Year(dateDebutSejour) = Year(dMoisCible) And Month(dateDebutSejour) = Month(dMoisCible) Then
                                
                                ' E. Somme
                                If IsNumeric(vDataResas(k, colMontant)) Then
                                    montantTotal = montantTotal + CDbl(vDataResas(k, colMontant))
                                End If
                            End If
                        End If
                    End If
                End If
            Next k
            
            ' Écriture du résultat
            loProj.DataBodyRange(i, j + 1).value = montantTotal
            
        Next j
    Next i
    
    'On écrit l'intitulé des colonnes pour mieux s'y retrouver.
    For i = 1 To 5
           loProj.DataBodyRange(0, i + 1).value = Format(DateAdd("m", i, Date), "mmmm yyyy")
           Next i
    

FinProgramme:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Erreur d'exécution : " & Err.Description, vbCritical
    Resume FinProgramme
End Sub




Sub CalculDelais()
    Init
    '-------------------------------------------------------------------------
   'Cette procédure calcule les montants par délais de réservation
    '-------------------------------------------------------------------------
   Dim T As Variant
   T = Range("ListeRésas")
   
   Dim D As Variant
   D = Range("DelaisReservations")
   
   Dim i, j, n, delai As Long
   
   'On met le tableau à zéro
   For i = 2 To 3
        For j = 1 To 8
            D(j, i) = 0
        Next j
    Next i
   
   Dim dateDebut, DateDebut2, dateDebut3 As Date
   dateDebut = DateAdd("d", -365, Date)
   dateDebut3 = DateAdd("d", -365, Date)
   
   DateDebut2 = DateAdd("d", -730, Date)
   
   
'On boucle sur les réservations N
For i = 1 To UBound(T)
    'On vérifie si la réservation a été prise il y a moins d'un an
    If T(i, idxResas("Booking_date")) >= dateDebut And (T(i, idxResas("Location")) = "Apollinaire" Or T(i, idxResas("Location")) = "Maury") Then
        n = n + 1
        delai = T(i, idxResas("Date Début")) - T(i, idxResas("Booking_date"))
        D(1, 2) = D(1, 2) + delai
        D(8, 2) = D(8, 2) + T(i, idxResas("Versement"))
        
        Select Case delai
            Case 0 To 7: D(2, 2) = D(2, 2) + T(i, idxResas("Versement"))
            Case 8 To 14: D(3, 2) = D(3, 2) + T(i, idxResas("Versement"))
            Case 15 To 30: D(4, 2) = D(4, 2) + T(i, idxResas("Versement"))
            Case 31 To 60: D(5, 2) = D(5, 2) + T(i, idxResas("Versement"))
            Case 61 To 180: D(6, 2) = D(6, 2) + T(i, idxResas("Versement"))
            Case Else: D(7, 2) = D(7, 2) + T(i, idxResas("Versement"))
        End Select
    End If
Next i
If n > 0 Then D(1, 2) = D(1, 2) / n

'On boucle sur les réservations N-1
n = 0
For i = 1 To UBound(T)
    'On vérifie si la réservation a été prise il y a moins d'un an
    If T(i, idxResas("Booking_date")) >= DateDebut2 And T(i, idxResas("Booking_date")) < dateDebut3 And (T(i, idxResas("Location")) = "Apollinaire" Or T(i, idxResas("Location")) = "Maury") Then
        n = n + 1
        delai = T(i, idxResas("Date Début")) - T(i, idxResas("Booking_date"))
        D(1, 3) = D(1, 3) + delai
        D(8, 3) = D(8, 3) + T(i, idxResas("Versement"))
        
        Select Case delai
            Case 0 To 7: D(2, 3) = D(2, 3) + T(i, idxResas("Versement"))
            Case 8 To 14: D(3, 3) = D(3, 3) + T(i, idxResas("Versement"))
            Case 15 To 30: D(4, 3) = D(4, 3) + T(i, idxResas("Versement"))
            Case 31 To 60: D(5, 3) = D(5, 3) + T(i, idxResas("Versement"))
            Case 61 To 180: D(6, 3) = D(6, 3) + T(i, idxResas("Versement"))
            Case Else: D(7, 3) = D(7, 3) + T(i, idxResas("Versement"))
        End Select
    End If
Next i

'Le delai moyen
If n > 0 Then D(1, 3) = D(1, 3) / n

'On met à jour le tableau
Range("DelaisReservations").value = D

End Sub

Sub CalculPriseReservations()
    '------------------------------------------------------------
    'Permet de calculer les prises de réservations
    'par mois
    '------------------------------------------------------------
    
    Dim T As Variant
    Dim R As Variant
    Dim i, j, anDebut, anFin, An, Mois, idBookingDate, idVersement, idLogement As Integer
    
    R = Range("ListeRésas")
    T = Range("PriseReservation")
    idBookingDate = Range("ListeRésas").ListObject.ListColumns("booking_date").Index
    idVersement = Range("ListeRésas").ListObject.ListColumns("Versement").Index
    idLogement = Range("ListeRésas").ListObject.ListColumns("Location").Index
    
    '--------------------------------------------------------------------------------
    '1. On préparer les différents variables
    '--------------------------------------------------------------------------------
    anDebut = CInt(Range("PriseReservation").ListObject.ListColumns(2))
    anFin = CInt(Range("PriseReservation").ListObject.ListColumns(UBound(T, 2)))
    
    For i = 2 To UBound(T, 2)
        For j = 1 To 13
            T(j, i) = 0
        Next j
    Next i
    
    '--------------------------------------------------------------------------------
    '2. On lance les calculs
    '--------------------------------------------------------------------------------
    For i = 1 To UBound(R)
        If Year(R(i, idBookingDate)) <= anFin And Year(R(i, idBookingDate)) >= anDebut Then
            If R(i, idLogement) = "Apollinaire" Or R(i, idLogement) = "Maury" Then
                T(Month(R(i, idBookingDate)), -anDebut + Year(R(i, idBookingDate)) + 2) = _
                T(Month(R(i, idBookingDate)), -anDebut + Year(R(i, idBookingDate)) + 2) + _
                R(i, 10)
            End If
        End If
    Next i

    For i = 2 To UBound(T, 2)
        For j = 1 To 12
            T(13, i) = T(13, i) + T(j, i)
        Next j
    Next i

 
    
     '--------------------------------------------------------------------------------
    '3. On renvoie la table
    '--------------------------------------------------------------------------------
   Range("PriseReservation") = T
End Sub
Sub CalculProchainsPaiements()
'---------------------------------------------------------------------
'Cette procédure remplit le tableau de la page d'accueil
'---------------------------------------------------------------------
    Dim T, PP As Variant
    
    T = Range("ListeRésas")
    Range("ProchainsPaiements").ListObject.DataBodyRange.ClearContents
    ReDim PP(1 To 20, 1 To 4)
    
    Dim R, n As Long
    
    '---------------------------------------------------------------------
    'On parcourt T pour calculer les prochains paiements
    '---------------------------------------------------------------------
    Dim depart As Long
        Init
    For R = UBound(T) To 1 Step -1
        If T(R, idxResas("Payé")) = "" Or (Date - T(R, idxResas("Date Début"))) < 7 Then
            n = n + 1
            'On met à jour T
             PP(n, 2) = T(R, 2)
             Select Case PP(n, 2)
                Case "Booking":
                    PP(n, 1) = T(R, 3) + 2 + T(R, 4)
                    PP(n, 3) = T(R, 10) + T(R, 9)
                Case "Airbnb":
                    PP(n, 1) = T(R, 3) + 2
                    PP(n, 3) = T(R, 10)
            End Select
            
            'On gère les week ends
            Select Case Weekday(PP(n, 1))
                Case 7: PP(n, 1) = PP(n, 1) + 2
                Case 1: PP(n, 1) = PP(n, 1) + 1
            End Select
                
            If n > 1 Then
                PP(n, 4) = PP(n - 1, 4) + PP(n, 3)
            Else
                PP(n, 4) = PP(n, 3)
            End If
            
            If n = 20 Then Exit For
        End If
    Next R
    
    'On remplit le tableau
    Range("ProchainsPaiements") = PP
End Sub

Sub MAJTravaux()
    '---------------------------------------------------------
    ' Cette procédure recalcule les tableaux après
    ' Modification des réservations
    '---------------------------------------------------------
    
   
    
    StatistiquesNotations
    CalculMenage (180)
    CalculMargeBrute
    CalculAnnulations
    
    CalculPriseReservations
    CalculProchainsPaiements
    CalculDelais
    
     majRecapitulatif
     CalculCalendrier
    
End Sub




Function GetMinMax(arrData As Variant, Logement As Long, Mois, annee) As String
    Dim vMin As Currency, vMax As Currency
    
    vMin = arrData(Logement, Mois, annee, 1)
    vMax = arrData(Logement, Mois, annee, 2)
    If vMin >= 0 Then
        GetMinMax = "(" & CStr(Int(vMin)) & " / " & CStr(Int(vMax)) & " €)"
    Else
        GetMinMax = ""
    End If
End Function



Function CalculMinMax() As Variant
    Dim t0 As Long
    t0 = ChronoStart()
    
    Dim wsResa As Worksheet, wsParam As Worksheet
    Dim tblResa As ListObject, tblLog As ListObject
    Dim arrResult() As Currency
    Dim nbLogements As Long, iLog As Long
    Dim annee As Long, Mois As Long
    Dim R As ListRow
    
    Dim idxColLogement As Long, idxColPrix As Long, idxColDate As Long
    Dim logName As String, prixNuit As Variant, dateDeb As Date
    Dim idxLog As Long
    
    Dim CUR_MAX_SENTINEL As Currency, CUR_MIN_SENTINEL As Currency
    CUR_MAX_SENTINEL = CCur(10000000000#)
    CUR_MIN_SENTINEL = CCur(-10000000000#)
    
       Dim dLog As Object
      Set dLog = BuildIndexCache("Logements", 1)
    
    '--- Feuilles et tables
    Set wsResa = ThisWorkbook.Worksheets("Réservations")
    Set wsParam = ThisWorkbook.Worksheets("Paramètres")
    Set tblResa = wsResa.ListObjects("listeRésas")
    Set tblLog = wsParam.ListObjects("Logements")
    
    nbLogements = tblLog.ListRows.Count
    
    '--- Indices de colonnes (noms supposés)
    idxColLogement = tblResa.ListColumns("Location").Index
    idxColPrix = tblResa.ListColumns("Nuitée").Index
    idxColDate = tblResa.ListColumns("Date Début").Index
    
    ' Dimensions : (logement 1..n, mois 0..12, année 2023..2030, minmax 1=min 2=max)
    ReDim arrResult(1 To nbLogements, 0 To 12, 2023 To 2030, 1 To 2)
    
    '--- Initialisation min/max
    For iLog = 1 To nbLogements
        For annee = 2023 To 2030
            For Mois = 0 To 12
                arrResult(iLog, Mois, annee, 1) = CUR_MAX_SENTINEL ' min init
                arrResult(iLog, Mois, annee, 2) = CUR_MIN_SENTINEL ' max init
            Next Mois
        Next annee
    Next iLog
    
    '--- Parcours des réservations
    For Each R In tblResa.ListRows
        logName = CStr(R.Range.Columns(idxColLogement).value)
        prixNuit = R.Range.Columns(idxColPrix).value
        dateDeb = R.Range.Columns(idxColDate).value
        
        If IsNumeric(prixNuit) And Not IsEmpty(prixNuit) And prixNuit > 0 Then
            annee = Year(dateDeb)
            If annee >= 2023 And annee <= 2030 Then
                Mois = Month(dateDeb)
                
                ' Trouver l’index du logement via la helper fournie
                idxLog = dLog(logName)
                If idxLog > 0 And idxLog <= nbLogements Then
                    '--- Min/Max mensuels
                    If CCur(prixNuit) < arrResult(idxLog, Mois, annee, 1) Then
                        arrResult(idxLog, Mois, annee, 1) = CCur(prixNuit)
                    End If
                    If CCur(prixNuit) > arrResult(idxLog, Mois, annee, 2) Then
                        arrResult(idxLog, Mois, annee, 2) = CCur(prixNuit)
                    End If
                    
                    '--- Min/Max annuels (mois = 0)
                    If CCur(prixNuit) < arrResult(idxLog, 0, annee, 1) Then
                        arrResult(idxLog, 0, annee, 1) = CCur(prixNuit)
                    End If
                    If CCur(prixNuit) > arrResult(idxLog, 0, annee, 2) Then
                        arrResult(idxLog, 0, annee, 2) = CCur(prixNuit)
                    End If
                End If
            End If
        End If
    Next R
    
    '--- Remplacer les sentinelles par 0 si aucune donnée
    For iLog = 1 To nbLogements
        For annee = 2023 To 2030
            For Mois = 0 To 12
                If arrResult(iLog, Mois, annee, 1) = CUR_MAX_SENTINEL Then
                    arrResult(iLog, Mois, annee, 1) = 0
                End If
                If arrResult(iLog, Mois, annee, 2) = CUR_MIN_SENTINEL Then
                    arrResult(iLog, Mois, annee, 2) = 0
                End If
            Next Mois
        Next annee
    Next iLog
    
    CalculMinMax = arrResult
    Call ChronoStop(t0, "CalculMinMax")
End Function



Sub CalculCalendrier()
    Dim t0 As Long
    t0 = ChronoStart()
    BeginAppState
    
    '==============================================================
    'On calcule le tableau en le remettant à zéro puis en remplissant
    '==============================================================
    '1. On récupère les données pour l'année en cours et l'année suivante
    '==============================================================
    Dim anDebut As Integer
    Dim anFin As Integer
    Dim MatriceOccupation As Variant
    Dim MatricePrix As Variant
    Dim tableauReservations As Variant
    Dim TableauLogements As Variant
    Dim i As Long: Dim j As Long: Dim k As Long: Dim R As Long
    Dim Occupe As Boolean
    
    anDebut = Year(Now)
    anFin = anDebut + 1
    MatriceOccupation = CalculJour(anDebut, anFin)
    MatricePrix = GetTablePrix
    tableauReservations = Feuil10.Range("ListeRésas").ListObject.DataBodyRange.value

    Dim Ticker As Variant
    ReDim Ticker(dicLog.Count)
    For i = 1 To UBound(Ticker): Ticker(i) = "x":     Next i
   
    
    '==============================================================
    '2. On efface les données affichées
    '==============================================================
    Dim TableauCalendrier
    Dim ContenuCalendrier
    Dim wsCal As Worksheet
    
    TableauCalendrier = "TableauCalendrier"
    ContenuCalendrier = "ContenuCalendrier"
    Set wsCal = Sheets("Calendrier")
    
    wsCal.Range(TableauCalendrier).ClearContents
    wsCal.Range(TableauCalendrier).Borders.LineStyle = xlLineStyleNone
    wsCal.Range(ContenuCalendrier).Font.size = 16
    wsCal.Range(ContenuCalendrier).Font.Name = "WingDings"
    wsCal.Range(ContenuCalendrier).ShrinkToFit = False

    'On récupère le tableau dans une table
    Dim Resultat As Variant
    Resultat = wsCal.Range(TableauCalendrier).value
    
    '==============================================================
    '3. On calcule les données
    '==============================================================
    Dim dateDebut As Long          'Date de début de l'affichage
    Dim dateFin As Long             'Date de fin de l'affichage
    
    dateDebut = DateSerial(anDebut, Month(Now), 1)
    dateFin = DateAdd("yyyy", 1, dateDebut) - 1
    
    Dim AffDate As String
    Dim PosColMois(12)
    Dim PosColFinMois(12)
    Dim posLundi(60) As Integer
    Dim nbLundis As Integer
    Dim idMois As Integer
    Dim col As Integer              'La colonne dans la tableResultat
    
    idMois = 0
    
    'On lance la boucle pour remplir toutes les données d'occupation
    '-----------------------------------------------------
    For i = dateDebut To dateFin
        
        col = i - dateDebut + 1     'On calcule quelle est la colonne concernée
        
        '3.1 on regarde si c'est le début du mois
        '==================================================
        If Day(i) = 1 Then
            '3.1.1 on met à jour le nom du mois dans la première ligne
            Resultat(1, col) = "'" + Format(i, "mmm yyyy")
            idMois = idMois + 1                                 'On affiche un mois de plus
            PosColMois(idMois) = LettresColonne(col + 1)    'On mémorise la position du premier jour du mois
            PosColFinMois(idMois) = LettresColonne(col)     'C'est la position d'avant
        End If
        
        '3.2 on met à jour le jour dans le titre du tableau
        '==================================================
        Resultat(2, col) = Day(i)
        
        If Weekday(i, vbMonday) = 1 Then
            'On met à jour la table des semaines
            nbLundis = nbLundis + 1
            posLundi(nbLundis) = col + 1
        End If
        
        
        '3.3 On fait défiler les logements
        '==================================================
        For j = 1 To dicLog.Count
            '3.3.1 On regarde s'il y a une réservation dans une des sources
            Occupe = False
            For k = 1 To dicSrc.Count
                If MatriceOccupation(i, dicLogA.Items()(j - 1), k) <> 0 Then
                    Occupe = True
                    If dicSrc.Keys()(k - 1) = "HomeExchange" Then
                        Ticker(j) = "n"
                    Else
                        Ticker(j) = "x"
                    End If
                    Exit For
                End If
            Next k
            
            If Occupe Then
                'On regarde s'il faut faire changer le ticker
                For R = 1 To UBound(tableauReservations, 1)
                    If tableauReservations(R, 1) = ListeLogements(j, 1) And tableauReservations(R, 3) = i And dicSrc.Keys()(k - 1) <> "HomeExchange" Then
                        If Ticker(j) = "x" Then Ticker(j) = "y" Else Ticker(j) = "x"
                        Exit For
                    End If
                Next R
                Resultat(j + 2, col) = Ticker(j)
            Else
                If i >= LBound(MatricePrix, 2) And i <= UBound(MatricePrix, 2) Then Resultat(j + 2, col) = MatricePrix(dicLogA.Items()(j - 1) - 1, i)
            End If
        Next j
    Next i

    '==============================================================
    '4. On peuple le tableau
    '==============================================================
    wsCal.Range(TableauCalendrier).value = Resultat

    '==============================================================
    '5. On effectue les bordures
    '==============================================================
    wsCal.Range("B2:NC2").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    wsCal.Range("B3:NC3").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    wsCal.Range("B3" + ":NC" + CStr(3 + dicLog.Count)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    
    'Pour les séparateurs de semaine
    Dim plage As Range
    Dim ws As Worksheet
    Dim iLundi As Integer
    Dim iCol As Integer
    
    Set ws = wsCal
    
    For iLundi = 1 To nbLundis
        If plage Is Nothing Then
            For iCol = 3 To dicLog.Count
                Set plage = ws.Cells(iCol, posLundi(iLundi))
            Next iCol
        Else
            For iCol = 3 To dicLog.Count + 3
                Set plage = Union(plage, ws.Cells(iCol, posLundi(iLundi)))
            Next iCol
        End If
    Next iLundi
    
    ' Vérification que la plage est bien créée
    If Not plage Is Nothing Then
            plage.Borders(xlEdgeLeft).LineStyle = xlDash
            plage.Borders(xlEdgeLeft).Weight = xlThin
            plage.Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    End If
    
    'Pour les séparateurs des mois
    For i = 1 To idMois - 1
        wsCal.Range(PosColMois(i) + "2" + ":" + PosColFinMois(i + 1) + CStr(3 + dicLog.Count)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    Next i

    '==============================================================
    '6. On met à jour les polices pour les prix
    '==============================================================
    Dim cell As Range
    Dim PlageNumerique As Range
    Set PlageNumerique = Sheets("Calendrier").Range("ND4")
    
    '6.1 On sélectionne les cellules concernées
    For Each cell In Sheets("Calendrier").Range("B4:NC" + CStr(3 + dicLog.Count))
        If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
            ' Ajoute la cellule à plageNumerique en utilisant Union
            Set PlageNumerique = Union(PlageNumerique, cell)
        End If
    Next cell
    
    '6.2 On fait la mise en forme sur les cellules qui nous intéressent
    With PlageNumerique
        .Font.Name = "Luciole"
        .Font.size = 6
        .ShrinkToFit = True
    End With
    EndAppState
    Call ChronoStop(t0, "CalculCalendrier")
    
    
    '==============================================================
    '6. On met à jour la liste des logements A
    '==============================================================
    Dim lActifs As Variant
    Feuil6.Range("TitreLogementsA").ClearContents
    lActifs = Feuil6.Range("TitreLogementsA")
    j = 0
    Dim Key0 As Variant
    For Each Key0 In dicLog
        j = j + 1
        lActifs(j, 1) = Key0
    Next Key0
    Feuil6.Range("TitreLogementsA").value = lActifs
    
End Sub

Function CalculLogementsActifs()
    '----------------------------------------------------------------
    'Cette fonction permet de réduire le nombre de logements
    'analysés en fonction de ceux qui sont actifs
    '----------------------------------------------------------------
    Dim Logements As Variant
    Dim n As Integer
    Logements = Range("Logements")
     
    Dim i As Integer
    For i = 1 To UBound(Logements)
        If Logements(i, 8) Then n = n + 1
    Next i
    
    CalculLogementsActifs = n
End Function

Sub CalculMargeBrute()
    Dim t0 As Long
    t0 = ChronoStart()
    '----------------------------------------------------------------
    'Cette procédure met à jour le tableau de marge brute
    'de la page d'accueil
    '---------------------------------------------------------------
    Dim tableauMB As Variant
    Dim tableauReservations As Variant
    Dim tableauCalcul() As Variant
    Dim anDebut, anFin, idLogement As Integer
    
    'Les tableaux de calcul et affichage
    Init
    tableauMB = Range("MBLofts")
    anDebut = tableauMB(1, 1)
    anFin = tableauMB(UBound(tableauMB), 1)
    ReDim tableauCalcul(anDebut To anFin, 4 * dicSrc.Count)
    
    'La base de travail
    tableauReservations = Range("ListeRésas")
    'Les logements concernés
    
    Dim i, j As Long
    Dim annee As Long
    Dim idChannel As String
    Dim numChannel As Integer
    
    '----------------------------------------------------------------
    ' 1. On lance le calcul
    '---------------------------------------------------------------
    For i = 1 To UBound(tableauReservations)
        'On récupère l'année de la réservation
        annee = Year(tableauReservations(i, 3))
        
        If annee >= anDebut And annee <= anFin Then
            idLogement = 0
            
            idChannel = 0
            'On récupère la source
            Dim cleLogement As Variant
            cleLogement = tableauReservations(i, 2)
    
            If cleLogement <> "" And dicSrc.Exists(cleLogement) Then
                idChannel = dicSrc(tableauReservations(i, 2))
            End If
           
            
            'On récupère le logement
            cleLogement = tableauReservations(i, 1)
    
            If cleLogement <> "" And dicLogs.Exists(cleLogement) Then
                idLogement = dicLogs(tableauReservations(i, 1))
            End If
            idLogement = dicLogs(tableauReservations(i, 1))
            
            If idChannel <> 0 And idLogement <> 0 Then
                idChannel = 2 * idChannel - 2
                'On met à jour les totaux dans le tableau
                tableauCalcul(annee, idChannel + 1) = tableauCalcul(annee, idChannel + 1) + CCur(tableauReservations(i, 8)) + CCur(tableauReservations(i, 9)) + CCur(tableauReservations(i, 10))
                tableauCalcul(annee, idChannel + 2) = tableauCalcul(annee, idChannel + 2) + CCur(tableauReservations(i, 10))
            End If
        End If
        
    Next i
    
    '----------------------------------------------------------------
    ' 2. On calcul le tableauMB
    '---------------------------------------------------------------
    For i = anDebut To anFin
        For j = 1 To dicSrc.Count
            If tableauCalcul(i, 2 * j - 1) <> 0 Then
                tableauMB(i - anDebut + 1, j + 1) = tableauCalcul(i, 2 * j) / tableauCalcul(i, 2 * j - 1)
            Else
                tableauMB(i - anDebut + 1, j + 1) = ""
            End If
        Next j
    Next i
    
    Range("MBLofts") = tableauMB
    
    Call ChronoStop(t0, "CalculMargeBrute")
End Sub

Sub CalculMenage(duree As Long)
    Dim t0 As Long: t0 = ChronoStart()
    'BeginAppState
    
    Dim tbl As Variant
    Dim res() As Variant
    Dim i As Long, n As Long
    Dim datej As Long
    Dim nbNuits As Long, dateDebut As Long, dateMenage As Long
    
    datej = CLng(Date)
    
    '--- Charger toute la plage ListeRésas en mémoire
    tbl = Range("ListeRésas").value
    
    '--- Préparer tableau résultat (même nb de lignes max que source)
    ReDim res(1 To 40, 1 To 2)
 
    For i = UBound(tbl, 1) To 1 Step -1
        nbNuits = CLng(tbl(i, 4))     ' <-- colonne Nb Nuits
        dateDebut = CLng(tbl(i, 3))   ' <-- colonne Date Début
        dateMenage = nbNuits + dateDebut
        
        If (dateMenage >= datej) And (dateMenage - duree <= datej) Then
            If n < 40 Then
                n = n + 1
                res(n, 1) = tbl(i, 1)  ' <-- colonne Location
                res(n, 2) = CDate(dateMenage)
            Else
                Exit For
            End If
            
        End If
    Next i
   
    '--- Ajuster taille réelle du résultat
    If n > 0 Then
        'TableToTableau res, "Menages", n
        Range("Menages") = res
        With Range("Menages").ListObject.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("Menages").ListObject.ListColumns(2).DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .header = xlYes
            .Apply
        End With
    Else
        ' tableau vide : tu peux effacer "Menages" si nécessaire
        Erase res
        TableToTableau res, "Menages"
    End If
    
    Call ChronoStop(t0, "CalculMenage")
    'EndAppState
End Sub

Sub CA12Mois()
    Dim t0 As Long
    t0 = ChronoStart()
    BeginAppState
    Init True
    '---------------------------------------------------------
    'Permet de trouver les revenus sur les douze derniers mois glissants
    '---------------------------------------------------------
    '1. On retrouve les CA par jours et par logements
    Dim TableauCA12Mois() As Variant
    Dim dateDebut, dateFin As Date
    Dim i, D, s As Long
    Dim TableauJour As Variant
    
    dateDebut = DateAdd("yyyy", -1, Date)   'Il y a un an
    dateFin = DateAdd("d", -1, Date)            'Hier
    TableauJour = CalculJour(Year(dateDebut), Year(dateFin))
    
    '2. On récupère les totaux entre les dates cibles
    ReDim TableauCA12Mois(1 To dicLogs.Count, 1 To dicSrc.Count + 2)
    For i = 1 To dicLogs.Count
        For D = dateDebut To dateFin
            For s = 1 To dicSrc.Count
                TableauCA12Mois(i, s + 1) = TableauCA12Mois(i, s + 1) + TableauJour(D, i, s)
            Next s
        Next D
        TableauCA12Mois(i, 1) = ListeLogements(i, 1)
    Next i

    '2.1 On fait le total par appartement. Ubound(TableauJour,3) contient le nombre de channels
    For i = 1 To dicLogs.Count
        For s = 1 To dicSrc.Count
            TableauCA12Mois(i, dicSrc.Count + 2) = TableauCA12Mois(i, dicSrc.Count + 2) + TableauCA12Mois(i, s + 1)
        Next s
    Next i

    '3. On met à jour le tableau CA12mois
    Feuil11.Range("CA12Mois").ListObject.DataBodyRange.value = TableauCA12Mois
    
    '4. On complète l'historique
    Dim n As Long
    Dim Total As Currency
    n = Feuil11.Range("HistoriqueCA").ListObject.ListRows.Count
    If CDate(Feuil11.Range("HistoriqueCA[Date]")(1)) <> CDate(Date) Then
        Feuil11.Range("HistoriqueCA").ListObject.ListRows.Add 1
        n = n + 1
        Feuil11.Range("HistoriqueCA[Date]")(1) = Date
        Feuil11.Range("HistoriqueCA[Date]")(1).NumberFormat = "dd/mm/yyyy"
     
        'On fait le total pour les deux lofts seulement
        Total = 0
        For i = 1 To 2
            Total = Total + TableauCA12Mois(i, UBound(TableauJour, 3) + 2)
        Next i
    
        Feuil11.Range("HistoriqueCA[Lofts]")(1) = Total
     
        'On fait le total pour tous les logements
        Feuil11.Range("HistoriqueCA[CA 12M]")(1) = Feuil11.Range("CA12Mois").ListObject.ListColumns(dicSrc.Count + 2).Total
        
        'On met le total du portefeuille
        Total = 0
        
        'On indique ainsi le stock
        For i = 1 To 2
            Total = Total + Feuil11.Range("AVenir").Cells(i, dicSrc.Count + 3)
        Next i
        Feuil11.Range("HistoriqueCA[Portefeuille]")(1) = Total
        
        'On met ici le stock N-1
        Feuil11.Range("HistoriqueCA[N-1]")(1) = CalculStockAN(Date)
        
    End If
    EndAppState
    Call ChronoStop(t0, "CA12Mois")
End Sub

Function CalculStock(anDebut, anFin)
    Dim t0 As Long
    t0 = ChronoStart()
    '---------------------------------------------------------
    ' Cette procédure permet d'avoir un aperçu du stock brut dispo
    '---------------------------------------------------------
    '1. On initie les variables
    '---------------------------------------------------------
    Dim ValeurStock()
    Dim nbLogements As Integer
    Dim iLogement As Integer
    Dim TablePrixAirbnb As Variant
    
    nbLogements = Range("Logements").ListObject.ListRows.Count
    ReDim ValeurStock(anDebut To anFin, 12, nbLogements)
    TablePrixAirbnb = GetTablePrix
    
    '---------------------------------------------------------
    ' 2. On fait les cumuls pour les prix disponibles dans U
    '---------------------------------------------------------
    Dim iJourPrix As Long
    
    For iLogement = 0 To nbLogements - 1
        For iJourPrix = LBound(TablePrixAirbnb, 2) To UBound(TablePrixAirbnb, 2)
            If TablePrixAirbnb(iLogement, iJourPrix) <> 0 Then
                ValeurStock(Year(iJourPrix), Month(iJourPrix), iLogement) = _
                    ValeurStock(Year(iJourPrix), Month(iJourPrix), iLogement) + _
                    TablePrixAirbnb(iLogement, iJourPrix)
            End If
        Next iJourPrix
    Next iLogement
    
    CalculStock = ValeurStock
    Call ChronoStop(t0, "CalculStock")
End Function

Function CalculStockAN(D As Date) As Currency
    Dim t0 As Long
    t0 = ChronoStart()
    '---------------------------------------------
    'Cette fonction retourne le stock de réservations disponibles
    'A une date déterminée
    '---------------------------------------------
    Dim strSql As String
    Dim FilterDate As String
    
    FilterDate = "#" + Format(DateAdd("yyyy", -1, D), "yyyy-mm-dd") + "#"
   
   strSql = "SELECT sum(Versement) " _
                + "FROM " + plage("ListeRésas") + " " _
                + "WHERE (Location = 'Apollinaire' or Location ='Maury') " _
                + "AND Booking_Date < " + FilterDate + " " _
                + "AND [Date Début] > " + FilterDate
                
  CalculStockAN = ExecuteSQL(strSql)(0, 0)
    
                
    Call ChronoStop(t0, "CalculStockAn")
End Function

Sub majCalculChannel()
    '----------------------------------------------------
    'Cette procédure permet de comparer les performances des channels
    '----------------------------------------------------
    Dim t0 As Long
    t0 = ChronoStart()
    Dim dd: dd = timer * 1000
    '1. Récupération des données
    Dim T As Variant
    T = CalculMois(2024, 2026, True)
    
    '2. Préparation du tableau
    Dim iAnnee As Integer
    Dim iMois As Integer
    Dim iLogement As Integer: iLogement = Feuil16.ComboBox1.ListIndex + 1
    Dim iChannel As Integer
    Dim nbLogements As Integer: nbLogements = UBound(T, 3)
    Dim nbChannels As Integer: nbChannels = UBound(T, 5)
    Dim Compteur As Integer
    Dim jLogement
    
    If iLogement = nbLogements + 1 Then Compteur = 2 Else Compteur = nbLogements
    Dim U As Variant
    
    For iAnnee = 2024 To 2026
        ReDim U(1 To 13, 1 To 4 * nbChannels)
        For iMois = 1 To 12
            For iChannel = 1 To nbChannels
                'On différentie si c'est une somme ou 1 seul logement
                If iLogement <= nbLogements Then
                    U(iMois, 4 * (iChannel - 1) + 1) = T(iAnnee, iMois, iLogement, 2, iChannel)
                    U(iMois, 4 * (iChannel - 1) + 2) = T(iAnnee, iMois, iLogement, 1, iChannel)
                Else
                    For jLogement = 1 To Compteur
                        U(iMois, 4 * (iChannel - 1) + 1) = U(iMois, 4 * (iChannel - 1) + 1) + T(iAnnee, iMois, jLogement, 2, iChannel)
                        U(iMois, 4 * (iChannel - 1) + 2) = U(iMois, 4 * (iChannel - 1) + 2) + T(iAnnee, iMois, jLogement, 1, iChannel)
                    Next jLogement
                End If
                If U(iMois, 4 * (iChannel - 1) + 2) <> 0 Then
                    U(iMois, 4 * (iChannel - 1) + 3) = CCur(U(iMois, 4 * (iChannel - 1) + 1) / U(iMois, 4 * (iChannel - 1) + 2))
                End If
            Next iChannel
        Next iMois
        
        'On fait le total
        For iChannel = 1 To nbChannels
            For iMois = 1 To 12
                U(13, 4 * (iChannel - 1) + 1) = U(13, 4 * (iChannel - 1) + 1) + U(iMois, 4 * (iChannel - 1) + 1)
                U(13, 4 * (iChannel - 1) + 2) = U(13, 4 * (iChannel - 1) + 2) + U(iMois, 4 * (iChannel - 1) + 2)
            Next iMois
            If U(13, 4 * (iChannel - 1) + 2) <> 0 Then
                U(13, 4 * (iChannel - 1) + 3) = CCur(U(13, 4 * (iChannel - 1) + 1) / U(13, 4 * (iChannel - 1) + 2))
            End If
        Next iChannel
        
        
        'On calcule les pourcentages
        Dim Total As Currency
        For iMois = 1 To 13
            Total = 0
            For iChannel = 1 To nbChannels
                Total = Total + U(iMois, 4 * (iChannel - 1) + 1)
            Next iChannel
            If Total <> 0 Then
                For iChannel = 2 To nbChannels
                    U(iMois, 4 * (iChannel - 1) + 4) = U(iMois, 4 * (iChannel - 1) + 1) / Total
                Next iChannel
            End If
        Next iMois
       
        
        
        'On met à jour
        Dim Irange As String
        Irange = "b" + CStr(16 * (iAnnee - 2024) + 4) + ":I" + CStr(16 * (iAnnee - 2024) + 16)
        Feuil16.Range(Irange).value = U
        Call ChronoStop(t0, "MajCalculChannel")
    Next iAnnee
    
End Sub

Sub StatistiquesNotations()
    Dim t0 As Long
    t0 = ChronoStart()
    Init
    '-------------------------------------------
    'Cette procédure remplit le tableau des statistiques
    '-------------------------------------------
    '1. On fait les déclarations pour les tableaux utilisés
    '-------------------------------------------
    Dim tableauReservations
    
    tableauReservations = Feuil10.Range("ListeRésas").ListObject.DataBodyRange.value
    
    Dim Notes()
    ReDim Notes(dicLogs.Count, 2 * dicSrc.Count, 2)
    Dim iReservation As Integer
    Dim channel As Integer
    Dim iLogement As Integer
    Dim iColonne As Integer
    
     '-------------------------------------------
    '2. On recherche toutes les notes obtenues
    '-------------------------------------------
   For iReservation = 1 To UBound(tableauReservations, 1)
        If tableauReservations(iReservation, 13) <> 0 Then
            channel = 2 * dicSrc(tableauReservations(iReservation, 2)) - 1
            
            'On met à jour l'historique
            iLogement = dicLogs(tableauReservations(iReservation, 1))
            Notes(iLogement, channel, 1) = Notes(iLogement, channel, 1) + 1
            Notes(iLogement, channel, 2) = Notes(iLogement, channel, 2) + tableauReservations(iReservation, 13)
        
            'On regarde si cela date de moins de douze mois
            If Abs(DateDiff("m", tableauReservations(iReservation, 3), Now)) <= 12 Then
                Notes(iLogement, channel + 1, 1) = Notes(iLogement, channel + 1, 1) + 1
                Notes(iLogement, channel + 1, 2) = Notes(iLogement, channel + 1, 2) + tableauReservations(iReservation, 13)
            End If
        End If
    Next iReservation

    '-------------------------------------------
    '4. On met à jour le tableau
    '-------------------------------------------
    Dim TableauNotes As Variant
    
    Feuil11.Range("Notations").ClearContents
    TableauNotes = Feuil11.Range("Notations").value
    For iLogement = 1 To dicLogs.Count
        For iColonne = 1 To 2 * dicSrc.Count
            If Notes(iLogement, iColonne, 1) <> 0 Then
                TableauNotes(iLogement, iColonne) = FormatNumber(Notes(iLogement, iColonne, 2) / Notes(iLogement, iColonne, 1), 2) _
                    + " (" + CStr(Notes(iLogement, iColonne, 1)) + ")"
            End If
        Next iColonne
    Next iLogement

    Feuil11.Range("Notations").value = TableauNotes
    Call ChronoStop(t0, "StatistiquesNotation")
End Sub
Function CalculJour(Optional anDebut = 2023, Optional anFin = 2030)
    Dim t0 As Long
    t0 = ChronoStart()
    '==================================================
    '1. On vérifie les dates et on calcule les plages
    '==================================================
    If Not IsNumeric(anDebut) Or Not IsNumeric(anFin) Then
        MsgBox "Erreur dans la saisie des dates"
        CalculJour = ""
        Exit Function
    End If
    
    Dim dateDebut As Date
    Dim dateFin As Date
    
    dateDebut = CDate("01/01/" + CStr(anDebut))
    dateFin = CDate("31/12/" + CStr(anFin))
    
    '==================================================
    '2. On Crée la Table et les variables dont on aura besoin
    '==================================================
    Dim TableauJour()
    
    'On dimensionne la table qui va contenir
    'les montants des loyers
    'pour chaque jour
    'Chaque Logement  (Prix net et prix brut)
    'chaque source
    'Le nbLogements est doublé car on calcul le tarif reçu puis le tarif payé par le client qui sont évidemment différents
    Init True
    ReDim TableauJour(CLng(dateDebut) To CLng(dateFin), dicLogs.Count * 2, dicSrc.Count)
    
    '==================================================
    '3. On récupère les réservations dans le tableau
    '==================================================
    Dim TableauReservation
    TableauReservation = Feuil10.Range("ListeRésas").value
    Dim iSource As Integer
    Dim iLogement As Integer
    Dim iReservation As Long
    Dim iJour As Long
    Dim prixNuit As Currency
    Dim prixNuitClient As Currency
    
    
    '==================================================
    '4. On peuple le tableauJour
    '==================================================
     'On parcourt la table du premier au dernier réservation dans la table
    For iReservation = LBound(TableauReservation, 1) To UBound(TableauReservation, 1)
        '4.1 On récupère la position de l'appartement dans la liste et la sources
        '--- AJOUT DE SÉCURITÉ ---
        ' On vérifie que la cellule n'est pas vide pour éviter l'ajout d'une clé vide
        Dim cleLogement As Variant
        cleLogement = TableauReservation(iReservation, 1)
    
        If cleLogement <> "" And dicLogs.Exists(cleLogement) Then
            iLogement = dicLogs(TableauReservation(iReservation, 1))
        Else
            MsgBox "Un problème avec ce logement qui n'a pas de titre"
            Stop
        End If
        
        cleLogement = TableauReservation(iReservation, 2)
    
        If cleLogement <> "" And dicSrc.Exists(cleLogement) Then
            iSource = dicSrc(TableauReservation(iReservation, 2))
        Else
            MsgBox "Un problème avec ce logement qui n'a pas de titre"
            Stop
        End If
      
       
        '4.2 on fait une boucle avec les valeurs de dates
        For iJour = TableauReservation(iReservation, 3) To TableauReservation(iReservation, 3) + TableauReservation(iReservation, 4) - 1
            If iJour >= LBound(TableauJour, 1) And iJour <= UBound(TableauJour, 1) Then
                '4.2.0 On récupère le prix à la nuit versé
                prixNuit = CCur(TableauReservation(iReservation, 10) / TableauReservation(iReservation, 4))
            
                '4.2.1 On récupère le prix à la nuit payé
                prixNuitClient = CCur(TableauReservation(iReservation, 5))
                
                '4.2.2 On met à jour le tableau
                TableauJour(iJour, iLogement, iSource) = TableauJour(iJour, iLogement, iSource) + prixNuit
                TableauJour(iJour, iLogement + dicLogs.Count, iSource) = TableauJour(iJour, iLogement + dicLogs.Count, iSource) + prixNuitClient
            End If
        Next iJour
        
    Next iReservation
    
    CalculJour = TableauJour
    'MsgBox (Timer * 1000 - DD)
    Call ChronoStop(t0, "CalculJour")
End Function


Function CalculMois(Optional anDebut = 2023, Optional anFin = 2030, Optional DetailChannel = False)
    Dim t0 As Long
    t0 = ChronoStart()
    '===================================================
    '1. On controle la validité des dates
    '===================================================
    If Not IsNumeric(anDebut) Or Not IsNumeric(anFin) Or anFin < anDebut Then
        MsgBox "Erreur dans la saisie des dates"
        CalculMois = ""
        Exit Function
    End If
    
    '===================================================
    '2. On récupère les données par jour
    '===================================================
    Dim TableauJour As Variant
    TableauJour = CalculJour(anDebut, anFin)
    
    '===================================================
    '3. On dimensionne la matrice
    '===================================================
    Dim nbLogements As Integer
    Dim nbChannels As Integer
    nbChannels = UBound(Range("Sources").value)
   
    Dim TableauMois()
    nbLogements = Range("Logements").ListObject.ListRows.Count
    
    If DetailChannel Then
        ReDim TableauMois(anDebut To anFin, 12, nbLogements, 2, nbChannels)
    Else
        ReDim TableauMois(anDebut To anFin, 12, nbLogements * 2, 2)
    End If
    
    '===================================================
    '4. On fait défiler pour faire tous les calculs
    '===================================================
    Dim iAnnee As Integer
    Dim iMois As Integer
    Dim iLogement As Integer
    Dim iJour As Long
    Dim iChannel As Integer
    Dim TotalNet As Currency
    Dim totalBrut As Currency
   
    
    For iAnnee = anDebut To anFin                'On calcule pour chaque année
        For iMois = 1 To 12                      'Pour chaques mois de chaque année
            For iLogement = 1 To nbLogements     'Pour chaque logements
                For iJour = DateSerial(iAnnee, iMois, 1) To DateAdd("m", 1, DateSerial(iAnnee, iMois, 1)) - 1 'On fait défiler les jours
                    If DetailChannel Then
                        For iChannel = 1 To nbChannels
                            If TableauJour(iJour, iLogement, iChannel) <> 0 Then
                                TableauMois(iAnnee, iMois, iLogement, 1, iChannel) = TableauMois(iAnnee, iMois, iLogement, 1, iChannel) + 1
                                TableauMois(iAnnee, iMois, iLogement, 2, iChannel) = TableauMois(iAnnee, iMois, iLogement, 2, iChannel) + TableauJour(iJour, iLogement, iChannel)
                            End If
                        Next iChannel
                    Else
                        'On regarde si à cette date là, dans ce logements là il y a un prix de la chambre
                        TotalNet = 0
                        totalBrut = 0
                        For iChannel = 1 To nbChannels 'On additionne tous les channels de réservation
                            TotalNet = TotalNet + TableauJour(iJour, iLogement, iChannel)
                            totalBrut = totalBrut + TableauJour(iJour, iLogement + nbLogements, iChannel)
                        Next iChannel
                    
                        If TotalNet <> 0 Then
                            TableauMois(iAnnee, iMois, iLogement, 1) = TableauMois(iAnnee, iMois, iLogement, 1) + 1 'On indique le nombre de jours occupés
                            TableauMois(iAnnee, iMois, iLogement, 2) = TableauMois(iAnnee, iMois, iLogement, 2) + TotalNet
                            TableauMois(iAnnee, iMois, iLogement + nbLogements, 2) = TableauMois(iAnnee, iMois, iLogement + nbLogements, 2) + totalBrut
                        End If                   'Il y a eu une journée prise dans un channel
                    End If
                Next iJour                       'Boucle des jours du mois
            Next iLogement                       'Boucle des logements
        Next iMois                               'Boucle des mois
    Next iAnnee                                  'Boucle des années
    
    '===================================================
    '5. On renvoie le tableau mensuel
    '===================================================
    CalculMois = TableauMois
    Call ChronoStop(t0, "CalculMois")
End Function
Sub majStatut()
    Dim t0 As Long
    t0 = ChronoStart()
    BeginAppState
    Init
    '--------------------------------------------------------------------------------
    'Cette précédure permet de mettre à jour le statut des réservations
    'Refactorisation Janvier 2026
    '--------------------------------------------------------------------------------
    '1. On charge les tableaux
    '--------------------------------------------------------------------------------
    Dim tableauReservations As Variant
    Dim TableauAVEnir As Variant
    Dim TableauEnCours As Variant
    Dim codesReservations As Variant
      Dim dateJour, endD As Date
    
    Init            'On initialise pour le cas où cela n'aurait pas été fait
    
    tableauReservations = Range("ListeRésas").ListObject.DataBodyRange.value
    codesReservations = Range("CodeReservation").ListObject.DataBodyRange.value
    
    dateJour = Date
    
   '--------------------------------------------------------------------------------
     '2. On prépare les tableaux
    '--------------------------------------------------------------------------------
    Dim iLigne, nbLigne As Long
    Dim iColonne As Long
    ReDim TableauAVEnir(1 To dicLog.Count, 1 To dicSrc.Count + 3)
    ReDim TableauEnCours(1 To 5, 1 To 11)
    
    
    
    '2.1 On prépare les tableaux
    Dim k As Variant
    iLigne = 0
     For Each k In dicLog
            iLigne = iLigne + 1
            TableauAVEnir(iLigne, 1) = k
            TableauEnCours(iLigne, 1) = k
    Next
        
    
    '--------------------------------------------------------------------------------
    '3. On calcule les statut et les écrans d'accueil
    '--------------------------------------------------------------------------------
    For iLigne = 1 To UBound(tableauReservations, 1)
        endD = CLng(Int(tableauReservations(iLigne, 3))) + CLng(tableauReservations(iLigne, 4))
        If Int(tableauReservations(iLigne, 3)) > dateJour Then
            'C'est une réservation à venir on affiche son statut
            tableauReservations(iLigne, 15) = codesReservations(3, 2)
            
            'On cherche le logement correspondant
            iColonne = dicLog(tableauReservations(iLigne, 1))
            
             If TableauEnCours(iColonne, 4) = "" Or TableauEnCours(iColonne, 4) > tableauReservations(iLigne, 3) And iColonne <= dicLogs.Count Then
                 TableauEnCours(iColonne, 7) = tableauReservations(iLigne, 3)
                 TableauEnCours(iColonne, 8) = tableauReservations(iLigne, 18)
                 TableauEnCours(iColonne, 9) = tableauReservations(iLigne, 4)
                 TableauEnCours(iColonne, 11) = tableauReservations(iLigne, 10)
                 TableauEnCours(iColonne, 10) = Left(tableauReservations(iLigne, 2), 13)
            End If
                 
            TableauAVEnir(iColonne, 2) = CInt(TableauAVEnir(iColonne, 2)) + 1
            TableauAVEnir(iColonne, 3 + dicSrc.Count) = TableauAVEnir(iColonne, 3 + dicSrc.Count) + tableauReservations(iLigne, 10)
            
            TableauAVEnir(iColonne, 2 + dicSrc(tableauReservations(iLigne, 2))) = TableauAVEnir(iColonne, 2 + dicSrc(tableauReservations(iLigne, 2))) + 1
                        
            
        ElseIf endD <= dateJour Then
            'C'est une réservation qui est terminée
           tableauReservations(iLigne, 15) = codesReservations(5, 2)
            
        Else
            'C'est une réservation en cours
            tableauReservations(iLigne, 15) = codesReservations(4, 2)
            
            iColonne = dicLog(tableauReservations(iLigne, 1))
            TableauEnCours(iColonne, 2) = endD
             TableauEnCours(iColonne, 3) = tableauReservations(iLigne, 18)
             TableauEnCours(iColonne, 4) = tableauReservations(iLigne, 4)
             TableauEnCours(iColonne, 5) = tableauReservations(iLigne, 10)
             TableauEnCours(iColonne, 6) = tableauReservations(iLigne, 2)
        End If
    Next iLigne
    
    
  
    '**********************************************************
    'On met à jour les tableaux
    '**********************************************************
    Feuil11.Range("Avenir").ListObject.DataBodyRange.value = TableauAVEnir
    Feuil11.Range("EnCours").ListObject.DataBodyRange.value = TableauEnCours
    
    'On ne remet que la colonne à jour
    Dim n As Long: n = UBound(tableauReservations)
    Dim colStat() As Variant
    ReDim colStat(1 To n, 1 To 1)
    
    Dim i As Long
    For i = 1 To n
        colStat(i, 1) = tableauReservations(i, 15)
    Next
    
   Feuil10.Range("ListeRésas").ListObject.ListColumns(15).DataBodyRange.Value2 = colStat
    
    
    EndAppState
    Call ChronoStop(t0, "MajStatut")
End Sub





'----------------------------------------
' Renvoie l’index de colonne d’un en-tête exact (ligne 1), 0 si introuvable
'----------------------------------------
Private Function GetColByHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlByColumns, MatchCase:=False)
    If Not f Is Nothing Then
        GetColByHeader = f.Column
    Else
        GetColByHeader = 0
    End If
End Function

Sub majRecapitulatif(Optional anFin = 2030)
    Dim t0 As Long
    t0 = ChronoStart()
    'BeginAppState
    
    '----------------------------------------------------------------
    '1. On récupère les données mensuelles depuis la table des réservations
    '----------------------------------------------------------------
    Dim anDebut As Integer: anDebut = 2023
   
    
    Dim tableaRevenusMois As Variant
    Dim tableStockMois As Variant
    Dim tableauMinMax As Variant
    
    tableaRevenusMois = CalculMois(anDebut, anFin)
    tableStockMois = CalculStock(anDebut, anFin)
    tableauMinMax = CalculMinMax
     
    '---------------------------------------------------------------
    '2. On met à jour les intitulés des différents
    'Tableaux avec les données des années
    '---------------------------------------------------------------
    Dim iLogement As Long
    
    'Les noms des appartements
    For iLogement = 1 To dicLogB.Count
        Range("NomApt" + CStr(iLogement)) = ListeLogements(iLogement, 1)
    Next iLogement
    
    Dim nbAn  As Integer: nbAn = anFin - anDebut + 1
    Dim iAn As Integer
    
    
    '---------------------------------------------------------
    '3. On prépare les tableaux pour peupler
    'Les tableaux
    '---------------------------------------------------------
    Dim TableauResultat As Variant
    For iLogement = 1 To dicLogB.Count
        If dicLog(dicLogB.Keys()(iLogement - 1)) Then
            'On efface le contenu
            Range("RecapTableau" + CStr(iLogement)).ClearContents
            
            'On récupère le contenu du tableau pour le mettre à jour
            TableauResultat = Range("RecapTableau" + CStr(iLogement)).value
            
            'Boucle des dates
            Dim TotalCA, TotalBudget, TotalCout, TotalStock As Currency
            Dim TotalNuits As Long
            
            For iAn = anDebut To anFin                 'Boucle des lignes
                TotalCA = 0
                TotalBudget = 0
                TotalCout = 0
                TotalNuits = 0
                TotalStock = 0
                
                Dim iMois As Integer
                For iMois = 1 To 12                      'Boucle des mois
                    'On calcule les différentes valeurs qui nous intéressent
                    'Le CA et CA Payé
                    TableauResultat(1 + 4 * (iAn - anDebut), iMois) = tableaRevenusMois(iAn, iMois, iLogement, 2)
                   
                    'Le prix par nuit
                    If tableaRevenusMois(iAn, iMois, iLogement, 1) <> 0 Then
                        TableauResultat(3 + 4 * (iAn - anDebut), iMois) = _
                              CStr(Round(tableaRevenusMois(iAn, iMois, iLogement, 2) / tableaRevenusMois(iAn, iMois, iLogement, 1), 0)) _
                            + " / " _
                            + CStr(Round(tableaRevenusMois(iAn, iMois, iLogement + dicLogB.Count, 2) / tableaRevenusMois(iAn, iMois, iLogement, 1), 0)) _
                            + " €" + Chr(10) + GetMinMax(tableauMinMax, iLogement, iMois, iAn)
                    End If
                    
                    'Le taux de remplissage
                    If tableaRevenusMois(iAn, iMois, iLogement, 1) <> 0 Then
                        TableauResultat(4 + 4 * (iAn - anDebut), iMois) = tableaRevenusMois(iAn, iMois, iLogement, 1) / nbJoursMois(iMois, iAn)
                    End If
                    
                    'Le stock restant
                    If tableStockMois(iAn, iMois, iLogement) <> 0 Then
                        TableauResultat(4 + 4 * (iAn - anDebut), iMois) = Format(tableStockMois(iAn, iMois, iLogement), "(# ##0 €)") + " - " + CStr(CInt(TableauResultat(4 + 4 * (iAn - anDebut), iMois) * 100)) + "%"
                    End If
                   
                    'On calcul les totaux
                    TotalCA = TotalCA + tableaRevenusMois(iAn, iMois, iLogement, 2)
                    TotalCout = TotalCout + tableaRevenusMois(iAn, iMois, iLogement + dicLogB.Count, 2)
                    TotalNuits = TotalNuits + tableaRevenusMois(iAn, iMois, iLogement, 1)
                    TotalStock = TotalStock + tableStockMois(iAn, iMois, iLogement)
                Next iMois
                
                'On met à jour les totaux
                TableauResultat(1 + 4 * (iAn - anDebut), iMois) = TotalCA
                If TotalNuits <> 0 Then
                    TableauResultat(3 + 4 * (iAn - anDebut), iMois) = _
                                                       CStr(Round(TotalCA / TotalNuits, 0)) _
                                                     + " / " _
                                                     + CStr(Round(TotalCout / TotalNuits, 0)) _
                                                     + " €" + Chr(10) + GetMinMax(tableauMinMax, iLogement, 0, iAn)
                        
                    TableauResultat(4 + 4 * (iAn - anDebut), iMois) = TotalNuits / 365
                    
                    If TotalStock <> 0 Then
                        TableauResultat(4 + 4 * (iAn - anDebut), iMois) = Format(TotalStock, "(# ##0 €)") + " - " + CStr(CInt(TableauResultat(4 + 4 * (iAn - anDebut), iMois) * 100)) + " %"
                    End If
                End If
            Next iAn
            
            'On met à jour le budget 2024 à anfin
            Dim PosX, PosY, LigX
            For iAn = 2024 To anFin
                TotalBudget = 0
                'Colonne du budget à prendre
                'La colonne de la source
                PosY = iLogement * 2
                
                
                'La ligne de la destination
                LigX = 6 + (iAn - 2024) * 4
                
                For iMois = 1 To 12
                    TableauResultat(LigX, iMois) = Range("Budget" + CStr(iAn)).Cells(iMois, PosY).value
                    TotalBudget = TableauResultat(LigX, iMois) + TotalBudget
                Next iMois
                TableauResultat(LigX, 13) = TotalBudget
            Next iAn
            
            'On met à jour le tableau
            Range("RecapTableau" + CStr(iLogement)) = TableauResultat
        End If
    Next iLogement
    
    Call ChronoStop(t0, "MajRecapitulatif")
    EndAppState
End Sub


