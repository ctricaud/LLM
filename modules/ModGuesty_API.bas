Attribute VB_Name = "ModGuesty_API"

Option Explicit

'---- Constants

Private Const GUESTY_TOKEN_URL As String = "https://open-api.guesty.com/oauth2/token"
Private Const GUESTY_LISTINGS_URL As String = "https://open-api.guesty.com/v1/reservations?" _
            & "fields=money.payments.status%20money.payments.paidAt%20money.payments.amount%20status%20checkIn%20checkOut%20lastUpdatedAt%20createdAt%20nightsCount%20confirmationCode&" _
            & "sort=-checkIn&limit=100"
Private Const GUESTY_REVIEWS_URL As String = "https://open-api.guesty.com/v1/reviews?limit=100"
Private Const GUESTY_PRICES_URL As String = "https://open-api.guesty.com/v1/availability-pricing/api/calendar/listings"
Private Const GUESTY_RESERVATION_DETAIL As String = "https://open-api.guesty.com/v1/reservations/"

'-------------------------Les variables
Public dicSource, dicLogement As Object
   
Private Function LISTEID() As Variant
    LISTEID = Array("672c9288bd52280010d9bad8", "672c9265a689710012eb3c90", "672c92bd0f95e1001361a6a7")
End Function
Sub GuestyAddReservation(idReservation As Variant)
    '------------------------------------------------------------------
    'Ajoute la réservatio à listeRésas
    '--------------
    Dim dicResa As Object
    
    Set dicResa = GuestyGetReservation(idReservation)
    
    'On créée la table qui contient les éléments à insérer
    Dim T(1 To 26)
    
    'On insère les informations
    Dim Key
    For Each Key In dicResa
        If idxResas(Key) <> "" Then
            T(idxResas(Key)) = dicResa(Key)
        End If
    Next Key
    
    'On récupère la commission
    Dim L As Variant
    Dim Comm
    L = Feuil5.Range("Logements")
    Comm = CDbl(L(dicLog(T(1)), 5))
    
    '------------------------------------------------------------------
    '2. Mise à jour des champs
    '------------------------------------------------------------------
    T(idxResas("Versement")) = CCur((T(idxResas("Prix")) - T(idxResas("Ménage")) - T(idxResas("Frais channel"))) * (1 - Comm))
    T(idxResas("Frais Conciergerie")) = CCur(T(idxResas("Ménage")) + T(idxResas("Versement")) * Comm / (1 - Comm))
    T(idxResas("Nuitée")) = CCur(T(idxResas("Prix")) / T(idxResas("Nb Nuits")))
    
     T(idxResas("Date Début")) = Int(T(idxResas("Date Début")))
    T(idxResas("Booking_date")) = Int(T(idxResas("Booking_date")))
    
    '------------------------------------------------------------------
    '3. On insère la réservation
    '------------------------------------------------------------------
    Feuil10.Range("ListeRésas").ListObject.ListRows.Add 1
    Feuil10.Range("ListeRésas").ListObject.ListRows(1).Range.value = T
    
    

    '------------------------------------------------------------------
    '2. On met à jour les logs
    '------------------------------------------------------------------
    Dim texte As String
    
    texte = "Nouvelle réservation " + T(2) + " :" + Chr(10) _
        & T(1) & " arrivée le " & Format(T(3), "dd/mm/yyyy") & " pour " + T(4) + " nuits." & Chr(10) _
        & "Versement : " & T(10) & " €"
    log texte
    log ""
        
End Sub




Public Function GuestyGetReservation(ByVal reservationId As String) As Object

    '----------------------------------------------------------------------------
    'Permet de récupérer toutes les informations concernant les résas
    '----------------------------------------------------------------------------
     '----------------------------------------------------------------------------
    '1. On récupère les informations sur la réservation
    '----------------------------------------------------------------------------
    Dim token As String
    token = GetGuestyToken()
    
    Dim url As String
    url = GUESTY_RESERVATION_DETAIL & reservationId & "?fields=money.hostPayout%20guestStay.createdAt" _
        & "%20guest.phone%20guest.fullName%20numberOfGuests.numberOfAdults%20numberOfGuests.numberOfChildren%20numberOfGuests.numberOfInfants%20numberOfGuests.numberOfPets" _
        & "%20money.payments.fees.amount" _
        & "%20money.fareAccommodationAdjusted%20nightsCount%20checkIn%20money.fareCleaning%20money.hostServiceFee%20money.totalTaxes%20money.totalPaid"


    Dim http As Object ' WinHttp.WinHttpRequest.5.1
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & token
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 911, , "HTTP " & http.Status & " - " & http.responseText
    End If


    '---------------------------------------------------------------------------------
   '2. On créée les dictionnaires dont on a besoin
   '---------------------------------------------------------------------------------
   CreationDictionnaires
   
   '---------------------------------------------------------------------------------
    '3. On récupère les informations qui nous intéressent
   '---------------------------------------------------------------------------------
    Dim dic, ret As Object
    Set dic = ParseJson(http.responseText)
    
    Dim dicRet As Object
    Set dicRet = New Dictionary
    Dim Fees As Currency
    
    dicRet("Booking_date") = ISO8601ToDate(dic("obj.guestStay.createdAt"))
    dicRet("updated_At") = ISO8601ToDate(dic("obj.guestStay.updatedAt"))
    dicRet("Code réservation") = reservationId
    dicRet("status") = dic("obj.status")
    dicRet("phone") = "'" + dic("obj.guest.phone")
    dicRet("guest user") = dic("obj.guest.fullName")
    dicRet("lastName") = dic("obj.guest.lastName")
    dicRet("firstName") = dic("obj.guest.firstName")
    dicRet("guestEmail") = dic("obj.guest.email")
    dicRet("listing_id2") = dic("obj.listingId")
    dicRet("Location") = dicLogement(dic("obj.listingId"))
    dicRet("Source") = dicSource(dic("obj.integration.platform"))
    dicRet("Currency") = dic("obj.listing.prices.currency")
    dicRet("guestCount") = dic("obj.guestsCount")
    dicRet("adults") = dic("obj.numberOfGuests.numberOfAdults")
    dicRet("chilsdren") = dic("obj.numberOfGuests.numberOfChildren")
    dicRet("infants") = dic("obj.numberOfGuests.numberOfInfants")
    dicRet("pets") = dic("obj.numberOfGuests.numberOfPets")
    dicRet("Nb Nuits") = dic("obj.nightsCount")
    dicRet("Date Début") = ISO8601ToDate(dic("obj.checkIn"))
    dicRet("checkOut") = ISO8601ToDate(dic("obj.checkOut"))
     dicRet("PrixOriginal") = dic("obj.money.fareAccommodation")
     If dic("obj.money.payments(0).fees(0).amount") <> "" Then
        Fees = CCur(Replace(dic("obj.money.payments(0).fees(0).amount"), ".", ","))
    Else
        Fees = 0
    End If
    dicRet("Frais channel") = CCur(Replace(dic("obj.money.hostServiceFee"), ".", ",")) + Fees
    dicRet("Ménage") = CCur(Replace(dic("obj.money.fareCleaning"), ".", ","))
    dicRet("Prix") = CCur(Replace(dic("obj.money.fareAccommodationAdjusted"), ".", ",")) + CCur(dicRet("Ménage"))
   dicRet("Solde") = dic("obj.money.balanceDue")
    dicRet("fullyPaid") = dic("obj.money.isFullyPaid")

    
    Set GuestyGetReservation = dicRet

End Function
Sub CreationDictionnaires()
    CalculIdxResas
   Set dicSource = New Dictionary
   Set dicLogement = New Dictionary
     
    dicSource("airbnb2") = "Airbnb"
   dicSource("bookingCom") = "Booking"
   dicSource("manual") = "Hobe"
   dicSource("Booking.com") = "Booking"
   dicSource("homesVillasByMarriott") = "Marriott"
   
   dicLogement(LISTEID(0)) = "Apollinaire"
   dicLogement(LISTEID(1)) = "Maury"
   dicLogement(LISTEID(2)) = "Joséphine"
   
End Sub

Function GetGuestyToken_v0()
    Dim http As Object
    Dim url As String, postData As String
    Dim Response As String
    
    ' On vérifie s'il est nécessaire d'avoir un nouveau token
    If Now < Range("dateToken") Then
        GetGuestyToken = Range("lastToken")
        Exit Function
    End If
    
    url = GUESTY_TOKEN_URL
    
    ' 2. Corps de la requête encodée (form-urlencoded)
    postData = "grant_type=client_credentials" & _
               "&scope=open-api" & _
               "&client_id=" & URLEncode(GUESTY_CLIENT_ID) & _
               "&client_secret=" & URLEncode(GUESTY_CLIENT_SECRET)
    
    ' 3. Création objet WinHttp
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' 4. Envoi
    http.send postData
    
    ' 5. Réponse
    Response = http.responseText
    'Debug.Print "HTTP status:", http.Status
    'Debug.Print "Réponse JSON:", Response
    
    ' 6. Extraction access_token (rapide)
    Dim dic As Object
    
    Set dic = ParseJson(Response)
    Range("lastToken") = dic("obj.access_token")
    
    GetGuestyToken = Range("lastToken")
    Range("dateToken") = DateAdd("s", CLng(dic("obj.expires_in")), Now)
    
 
End Function
Function GetGuestyToken() As String
    ' Déclaration explicite des variables
    Dim http As Object
    Dim dic As Object
    Dim url As String, postData As String, responseBody As String
    Dim cacheDate As Variant
    
    ' --- 1. GESTION D'ERREUR ---
    On Error GoTo ErrHandler
    
    ' --- 2. VÉRIFICATION DU CACHE (Sécurisée) ---
    ' On utilise un Variant pour cacheDate pour éviter l'erreur "Type Mismatch" si la cellule est vide
    cacheDate = Feuil1.Range("dateToken").value
    
    If IsDate(cacheDate) Then
        If Now < CDate(cacheDate) Then
            ' Le token est encore valide, on le retourne directement
            GetGuestyToken = Feuil1.Range("lastToken").value
            Exit Function
        End If
    End If
    
    ' --- 3. PRÉPARATION DE LA REQUÊTE ---
    url = GUESTY_TOKEN_URL
    
    postData = "grant_type=client_credentials" & _
               "&scope=open-api" & _
               "&client_id=" & URLEncode(GUESTY_CLIENT_ID) & _
               "&client_secret=" & URLEncode(GUESTY_CLIENT_SECRET)
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' TIMEOUTS (Important !) : Resolve, Connect, Send, Receive (en ms)
    ' Ici : 5s pour connecter, 10s pour recevoir. Évite le gel d'Excel.
    http.setTimeouts 5000, 5000, 10000, 10000
    
    http.Open "POST", url, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' --- 4. ENVOI ET CONTRÔLE DE RÉCEPTION ---
    http.send postData
    
    ' Vérification du statut HTTP (200 = OK)
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, "GetGuestyToken", _
        "Erreur API (" & http.Status & "): " & http.responseText
    End If
    
    responseBody = http.responseText
    
    ' --- 5. PARSING ET STOCKAGE ---
    Set dic = ParseJson(responseBody)
    
    ' Note : Vérifiez bien la structure de votre JSON.
    ' Standard OAuth2 : dic("access_token") et dic("expires_in")
    ' Si votre ParseJson retourne une structure à plat, gardez votre syntaxe précédente.
    
    If dic Is Nothing Then Err.Raise vbObjectError + 2, "GetGuestyToken", "Erreur de parsing JSON"
    
    ' On vérifie que la clé existe avant de l'utiliser
    If Not dic.Exists("access_token") Then
        ' Fallback si votre parser utilise une autre clé (comme dans votre code original)
        If dic.Exists("obj.access_token") Then
             Feuil1.Range("lastToken").value = dic("obj.access_token")
             Feuil1.Range("dateToken").value = DateAdd("s", CLng(dic("obj.expires_in")), Now)
        Else
             Err.Raise vbObjectError + 3, "GetGuestyToken", "Token introuvable dans la réponse JSON"
        End If
    Else
        ' Cas Standard
        Feuil1.Range("lastToken").value = dic("access_token")
        ' On retire 60 secondes à la date d'expiration pour avoir une marge de sécurité
        Feuil1.Range("dateToken").value = DateAdd("s", CLng(dic("expires_in")) - 60, Now)
    End If
    
    GetGuestyToken = Range("lastToken").value
    
    ' Nettoyage mémoire
    Set http = Nothing
    Set dic = Nothing
    Exit Function

ErrHandler:
    ' En cas d'erreur, on affiche le problème dans la fenêtre Exécution et on retourne une chaine vide
    Debug.Print "ERREUR GetGuestyToken: " & Err.Description
    ' Optionnel : MsgBox "Impossible de récupérer le token Guesty : " & Err.Description, vbCritical
    GetGuestyToken = ""
    Set http = Nothing
    Set dic = Nothing
End Function
Public Sub GuestyGetReservations(Optional effaceLog = True)
    'Récupération des cent dernières réservations
    If effaceLog Then Range("logExtraction") = ""
   '--------------------------------------------------------------------------
    ' 1) Token (ta fonction existante)
    '--------------------------------------------------------------------------
   Dim token As String
    token = GetGuestyToken()
    
    '--------------------------------------------------------------------------
   ' 2) Appel API (100 dernières, tri décroissant)
    '--------------------------------------------------------------------------
   Dim url As String
    Dim i As Long, filterJson As String
    Dim filterEncoded As String
    
    '--- Construction dynamique du JSON ---
    Dim values As String

    For i = LBound(LISTEID) To UBound(LISTEID)
        values = values & """" & CStr(LISTEID(i)) & """"
        If i < UBound(LISTEID) Then values = values & ","
    Next i

    ' [{"operator":"$in","field":"listingId","value":["id1","id2"]},]
    'filterJson = "[{""operator"":""$in"",""field"":""listingId"",""value"":[" & values & "]},]"
filterJson = "[{""operator"":""$in"",""field"":""listingId"",""value"":[" & values & "]}, " & _
             "{""operator"":""$in"",""field"":""status"",""value"":[""confirmed""]}]"
    '--- Encodage URL du JSON ---
    filterEncoded = URLEncode(filterJson)
    
    '--- Requête ---
    token = GetGuestyToken() ' Ta fonction existante pour obtenir le token
    url = GUESTY_LISTINGS_URL & "&filters=" & filterEncoded
    
  
    Dim http As Object ' WinHttp.WinHttpRequest.5.1
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setTimeouts 15000, 15000, 30000, 30000
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & token
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 701, , "HTTP " & http.Status & " - " & http.responseText
    End If

    '--------------------------------------------------------------------------
   ' 3) Réponse JSON
   '--------------------------------------------------------------------------
    Dim resp As String
    resp = http.responseText
    Dim dir As Object
    Set dir = ParseJson(resp)
    
    '--------------------------------------------------------------------------
   ' 4) Mis en table du retour
   '--------------------------------------------------------------------------
   CreationDictionnaires
   
    Dim T As Variant
    Dim nbReservations As Long
    ReDim T(1 To 100, 1 To 28)
    
    For i = 0 To 99
        T(i + 1, idxResas("Source")) = dicSource(dir("obj.results(" & CStr(i) & ").integration.platform"))
        T(i + 1, idxResas("Statut")) = dir("obj.results(" & CStr(i) & ").status")
        T(i + 1, idxResas("Code réservation")) = dir("obj.results(" & CStr(i) & ")._id")
        T(i + 1, idxResas("Date Début")) = ISO8601ToDate(dir("obj.results(" & CStr(i) & ").checkIn"))
        T(i + 1, idxResas("Nb Nuits")) = dir("obj.results(" & CStr(i) & ").nightsCount")
        T(i + 1, idxResas("Location")) = dicLogement(dir("obj.results(" & CStr(i) & ").listingId"))
        
        '------------------------------------------------------------------------------------
        'On regarde les paiements eventuels pour la ligne i
        '-----------------------------------------------------------------------------------
        If dir("obj.results(" & CStr(i) & ").money.payments(0).status") = "SUCCEEDED" Then
            T(i + 1, idxResas("Versement")) = dir("obj.results(" & CStr(i) & ").money.payments(0).amount")
            T(i + 1, idxResas("Payé")) = "ü"
        End If
        If T(i + 1, 1) = "" Then Exit For
        nbReservations = nbReservations + 1
        
    Next i
    
    '--------------------------------------------------------------------------
   ' 5) On met à jour le tableau structuré ListeGuesty
   '--------------------------------------------------------------------------
   TableToTableau T, "ListeGuesty", nbReservations
   
   Feuil13.Range("ListeGuesty").ListObject.ListColumns(idxResas("Date Début")).DataBodyRange.NumberFormat = "m/d/yyyy"
   
    '--- Trier par DateDébut en ordre décroissant
        With Feuil13.Range("ListeGuestY").ListObject
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.ListColumns("Date Début").Range, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With .Sort
            .header = xlYes
            .Apply
        End With
    End With

    '--- Supprimer tous les critères de tri
    Feuil13.Range("ListeGuestY").ListObject.Sort.SortFields.Clear

     
    '--------------------------------------------------------------------------
   ' 6) On traite le tableau
   '--------------------------------------------------------------------------
   GuestyTraitementReservations
   
   '--------------------------------------------------------------------------
   ' 8) On remet à jour les autres feuilles
   '--------------------------------------------------------------------------
   If effaceLog Then MAJTravaux
   
End Sub

Sub GuestyGetReviews(Optional effaceLog = True)
    '----------------------------------------------------------------
    'Cette procédure récupère les dernières notations
    '----------------------------------------------------------------
    If effaceLog Then Range("logExtraction") = ""
    '--------------------------------------------------------------------------
    ' 1) Token (ta fonction existante)
    '--------------------------------------------------------------------------
   Dim token As String
    token = GetGuestyToken()
    
    'On construit le dictionnaire des réservations
    Dim D As Object
    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = vbTextCompare
    
    Dim nr As Long
    nr = Range("listeRésas").ListObject.DataBodyRange.rows.Count
    If nr > 0 Then
        Dim indexR
        indexR = Feuil10.Range("listeRésas").ListObject.ListColumns("Code réservation").DataBodyRange.value
        
        Dim r As Long
        For r = 1 To nr
            D(indexR(r, 1)) = r
        Next r
    End If
    
    '--------------------------------------------------------------------------
   ' 2) Appel API (20 dernières, tri décroissant)
    '--------------------------------------------------------------------------
    Dim i As Long, filterJson As String
     
    Dim url As String
    
    'On boucle sur les logements
    Dim iLog As Integer
    
    For iLog = LBound(LISTEID) To UBound(LISTEID)
        url = GUESTY_REVIEWS_URL & "&listingId=" + LISTEID(iLog)
      
        Dim http As Object ' WinHttp.WinHttpRequest.5.1
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "GET", url, False
        http.setTimeouts 15000, 15000, 30000, 30000
        http.setRequestHeader "Accept", "application/json"
        http.setRequestHeader "Authorization", "Bearer " & token
        http.send
    
        If http.Status < 200 Or http.Status >= 300 Then
            Err.Raise vbObjectError + 701, , "HTTP " & http.Status & " - " & http.responseText
        End If
    
        '--------------------------------------------------------------------------
       ' 3) Réponse JSON
       '--------------------------------------------------------------------------
        Dim resp As String
        resp = http.responseText
        Dim dir As Object
        Set dir = ParseJson(resp)
        
       Dim reservationId As String
        Dim note As Integer
        
        '--------------------------------------------------------------------------
       ' 4) On parcourt toutes les notes
       '--------------------------------------------------------------------------
        For i = 0 To CInt(dir("obj.limit")) - 1
            If dir("obj.data(" + CStr(i) + ").reservationId") <> "" Then
                'On récupère la note
                reservationId = dir("obj.data(" + CStr(i) + ").reservationId")
                If dir("obj.data(" + CStr(i) + ").rawReview.overall_rating") <> "" Then
                    note = dir("obj.data(" + CStr(i) + ").rawReview.overall_rating")
                Else
                    note = dir("obj.data(" + CStr(i) + ").rawReview.scoring.review_score")
                End If
                'Debug.Print reservationId + "-" + CStr(note)
                'On met à jour la note si on trouve la réservation et que la note n'est pas déjà mise
                If D(reservationId) <> "" Then
                    If Range("ListeRésas").ListObject.DataBodyRange(D(reservationId), 13).value = "" Then
                        '--------------------------------------------------------------------------
                        '4.2) On regarde si la note est mise
                        '--------------------------------------------------------------------------
                        Range("ListeRésas").ListObject.DataBodyRange(D(reservationId), 13).value = CInt(note)
                        
                        '--------------------------------------------------------------------------
                        '4.2.1)On met un message d'information
                        '--------------------------------------------------------------------------
                        log "Note de " + CStr(note)
                        log Range("ListeRésas").ListObject.DataBodyRange(D(reservationId), 1).value + " - " + Range("ListeRésas").ListObject.DataBodyRange(D(reservationId), 2).value + " séjour du " + Format(Range("ListeRésas").ListObject.DataBodyRange(D(reservationId), 3).value, "dd/mm/yyyy")
                     End If
                   End If
            End If
        Next i
   Next iLog
   
   If effaceLog Then MAJTravaux
   
End Sub

Sub GuestyPaiementsAirbnb(Optional effaceLog = True)
    If effaceLog Then Range("logExtraction") = ""
    '----------------------------------------------------
    'On traite les paiements reçus airbnb
    '----------------------------------------------------
    
    Dim idResa As Object
    Dim T As Variant
    
    T = Range("ListeRésas").value           'C'est la liste des réservations
    Set idResa = New Dictionary             'C'est la liste de toutes les réservations
    
    Dim i As Long
    For i = 1 To UBound(T)
        idResa(T(i, 14)) = i
     Next i
     
     Dim U As Variant
     U = Range("ListeGuesty").value     'C'est la liste des réservations Guesty
    
    'On parcourt toutes les réservations dans Guesty
    For i = 1 To UBound(U)
        'On vérifie si un paiement est enregistré dans Guesty mais pas dans listeRésas
        If idResa(U(i, 14)) <> "" Then      'La réservation a été trouvée
        'If idResa(U(i, 14)) = 37 Then Stop
            If U(i, 12) <> "" And T(idResa(U(i, 14)), 12) <> "ü" Then       'La réservation n'est pas encore notée comme payée
                
                'On gère ici les paiements multiples
                '---------------------------------------
                If T(idResa(U(i, 14)), 4) > 30 Then             'Plus de trente jours c'est une réservations payables en plusieurs fois
                    'C'est un paiement multiple
                    If Abs(U(i, 10) - T(idResa(U(i, 14)), 10)) < 1 Then         'La totalité du paiement est reçue
                        'On a atteint le total
                        T(idResa(U(i, 14)), 12) = "ü"
                        log "Paiement final de la réservation " + U(i, 2) + " - " + U(i, 1)
                        log CStr(U(i, 10)) + " € en date du " + Format(U(i, 3), "dd/mm/yyyy")
                    Else
                        'On met le montant payé
                        If T(idResa(U(i, 14)), 12) < U(i, 10) Then
                            T(idResa(U(i, 14)), 12) = U(i, 10)          'C'est le total payé qui est indiqué ici
                            log "Paiement partiel de la réservation " + U(i, 2) + " - " + U(i, 1)
                            log CStr(U(i, 10)) + " € en date du " + Format(U(i, 3), "dd/mm/yyyy")
                        End If
                    End If
                        
                Else
                    'C'est un oneshot avec un seul paiement
                    If Abs(U(i, 10) - T(idResa(U(i, 14)), 10)) < 1 Then
                        'On inscrit la case à cocher
                        T(idResa(U(i, 14)), 12) = "ü"
                        
                        'On met un message de log
                        log "Paiement réservation " + U(i, 2) + " - " + U(i, 1)
                        log CStr(U(i, 10)) + " € en date du " + Format(U(i, 3), "dd/mm/yyyy")
                    End If
                    
                End If              'Est ce que c'est un paiement en plusieurs fois
            End If              'Le paiement est il déjà enregistré dans la table de réservations
        End If              'La réservation est-elle trouvée dans la liste des réservations
    Next i              'Boucle de tous les enregistrements de listeGuesty
    
    'On met à jour le tableau
    Range("ListeRésas") = T
End Sub

Sub GuestyRemoveReservation(iReservation)
    '-------------------------------------------------------------------
    'On supprime la réservation qui n'a pas été trouvée dans Guesty
    '-------------------------------------------------------------------
    '1. On met le message
    '-------------------------------------------------------------------
    Dim T As Variant
    T = Range("ListeRésas").rows(iReservation)

    Dim texte As String
    
    texte = "Annulation réservation " + T(1, 2) + " :" + Chr(10) _
        & T(1, 1) & " arrivée le " & Format(T(1, 3), "dd/mm/yyyy") & " pour " + CStr(T(1, 4)) + " nuits." & Chr(10) _
        & "Versement : " & CStr(T(1, 10)) & " €"
    log texte
    log ""
    
    '-------------------------------------------------------------------
    '2. On supprime la ligne
    '-------------------------------------------------------------------
    Range("ListeRésas").ListObject.ListRows(iReservation).Delete
    
End Sub

Sub GuestyTraitementReservations()
'--------------------------------------------------------------------------
   ' Cette procédure permet de traiter les réservations chargées
   '--------------------------------------------------------------------------
    '1. On initialise listeRésa pour enlever les filtres et trier suivant la date de début
   '--------------------------------------------------------------------------
    TriListeResas
    Init
    
    '--------------------------------------------------------------------------
    '2. On recherche Guesty dans Listerésas
   '--------------------------------------------------------------------------
    CompareResasDansGuesty
    
    '--------------------------------------------------------------------------
    '3. On cherche ListeRésas dans listeGuesty
   '--------------------------------------------------------------------------
    CompareGuestyDansResas
    
    TriListeResas
End Sub
'--- Clé composite stable : logement|source|yyyymmdd|nbNuits
Private Function KeyOf(Logement As Variant, Source As Variant, D As Variant, nbNuits As Variant) As String
    KeyOf = CStr(Logement) & "|" & CStr(Source) & "|" & Format(CDate(D), "yyyymmdd") & "|" & CStr(nbNuits)
End Function

Sub CompareResasDansGuesty()
    '--------------------------------------------------------------------------------
    'Permet de comparer les listes issues de Guesty et de ListeRésas
    '--------------------------------------------------------------------------------
    '1. On construit le dictionnaire des réservations existantes
    '--------------------------------------------------------------------------------
    Dim loR As ListObject, log As ListObject
    Dim D As Object, r As Long, nr As Long, nG As Long
    Dim k As String, out(), col As ListColumn
    
    Set loR = Feuil10.Range("ListeRésas").ListObject
    Set log = Feuil13.Range("ListeGuesty").ListObject
    
    '--- Construire l'index des réservations existantes (ListeRésas)
    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = vbTextCompare
    
    nr = loR.DataBodyRange.rows.Count
    If nr > 0 Then
        Dim aLogR, aSrcR, aDateR, aNuitR
        aLogR = loR.ListColumns("Location").DataBodyRange.value
        aSrcR = loR.ListColumns("Source").DataBodyRange.value
        aDateR = loR.ListColumns("Date Début").DataBodyRange.value
        aNuitR = loR.ListColumns("Nb Nuits").DataBodyRange.value
        
        For r = 1 To nr
            k = KeyOf(aLogR(r, 1), aSrcR(r, 1), aDateR(r, 1), aNuitR(r, 1))
            D(k) = r
        Next r
    End If
    
    '--------------------------------------------------------------------------------
    '2. On vérifie l'existence des réservations de listeGuesty
    '--------------------------------------------------------------------------------
    '--- Vérifier chaque résa de ListeGuesty contre l'index
    nG = log.DataBodyRange.rows.Count
    If nG = 0 Then Exit Sub
    
    Dim gLog, gSrc, gDate, gNuit, gidReservation
    gLog = log.ListColumns("Location").DataBodyRange.value
    gSrc = log.ListColumns("Source").DataBodyRange.value
    gDate = log.ListColumns("Date Début").DataBodyRange.value
    gNuit = log.ListColumns("Nb Nuits").DataBodyRange.value
    gidReservation = log.ListColumns(idxResas("Code réservation")).DataBodyRange.value
    
    ReDim out(1 To nG, 1 To 1)
    For r = 1 To nG
        k = KeyOf(gLog(r, 1), gSrc(r, 1), gDate(r, 1), gNuit(r, 1))
        If Not (D.Exists(k)) Then
            GuestyAddReservation gidReservation(r, 1)
        End If
            
    Next r
    
End Sub

Sub CompareGuestyDansResas()
    '--------------------------------------------------------------------------------
    'Permet de comparer les listes issues de Guesty et de ListeRésas
    '--------------------------------------------------------------------------------
    '1. On construit le dictionnaire des réservations existantes
    '--------------------------------------------------------------------------------
    Dim loR As ListObject, log As ListObject
    Dim D As Object, r As Long, nr As Long, nG As Long
    Dim k As String, out(), col As ListColumn
    CreationDictionnaires
    
    
    Set loR = Feuil13.Range("ListeGuesty").ListObject
    Set log = Feuil10.Range("ListeRésas").ListObject
    
    '--- Construire l'index des réservations existantes (ListeRésas)
    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = vbTextCompare
    
    nr = loR.DataBodyRange.rows.Count
    If nr > 0 Then
        Dim aLogR, aSrcR, aDateR, aNuitR
        aLogR = loR.ListColumns("Location").DataBodyRange.value
        aSrcR = loR.ListColumns("Source").DataBodyRange.value
        aDateR = loR.ListColumns("Date Début").DataBodyRange.value
        aNuitR = loR.ListColumns("Nb Nuits").DataBodyRange.value
        
        For r = 1 To nr
            k = KeyOf(aLogR(r, 1), aSrcR(r, 1), aDateR(r, 1), aNuitR(r, 1))
            D(k) = r
        Next r
    End If
    
    '--------------------------------------------------------------------------------
    '2. On vérifie l'existence des réservations de listeRésas
    '--------------------------------------------------------------------------------
    '--- Vérifier chaque résa de ListeGuesty contre l'index
    nG = log.DataBodyRange.rows.Count
    If nG = 0 Then Exit Sub
    
    Dim gLog, gSrc, gDate, gNuit, gidReservation
    gLog = log.ListColumns("Location").DataBodyRange.value
    gSrc = log.ListColumns("Source").DataBodyRange.value
    gDate = log.ListColumns("Date Début").DataBodyRange.value
    gNuit = log.ListColumns("Nb Nuits").DataBodyRange.value
    gidReservation = log.ListColumns(idxResas("Code réservation")).DataBodyRange.value
    
    For r = 1 To nG
        k = KeyOf(gLog(r, 1), gSrc(r, 1), gDate(r, 1), gNuit(r, 1))
        If CLng(Now) - CLng(gDate(r, 1)) > 60 Then Exit For
        'If dicLogement(gLog(r, 1)) And gNuit(r, 1) > 0 Then
        If gNuit(r, 1) > 0 Then
            If Not (D.Exists(k)) And gSrc(r, 1) <> "HomeExchange" Then
                GuestyRemoveReservation r
            End If
        End If
    Next r
    
End Sub

Sub MajGuesty()
    Dim t0 As Long
    t0 = ChronoStart()
   '--------------------------------------------------------------
    'Cette procédure permet de tout mettre à jour
    '--------------------------------------------------------------
    Range("LogExtraction") = ""
    log Format(Now, "dd/mm/yyyy hh:nn")
    
    log "********** Réservations **********"
    GuestyGetReservations False
    
    log ""
    log "********** Reviews **********"
    GuestyGetReviews False
    
    log ""
    log "********** Paiements **********"
   GuestyPaiementsAirbnb False
    
    log ""
    log "********** Prix **********"
     GuestyGetPrixOnline
     
    MAJTravaux
    
   
    
    Call ChronoStop(t0, "MajGuesty")
End Sub
Sub TriListeResas()
    RAZFiltres "ListeRésas"
    With Range("ListeRésas").ListObject
        .ListColumns(idxResas("Date Début")).DataBodyRange.NumberFormat = "dd/mm/yyyy"
   
        '--- Trier par DateDébut en ordre décroissant
    
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.ListColumns("Date Début").Range, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With .Sort
            .header = xlYes
            .Apply
        End With
    End With
    
    '--- Supprimer tous les critères de tri
    Range("ListeRésas").ListObject.Sort.SortFields.Clear
End Sub

Function URLEncode(ByVal sText As String) As String
    Dim i As Long, sRes As String, sChar As String
    For i = 1 To Len(sText)
        sChar = Mid$(sText, i, 1)
        Select Case Asc(sChar)
            Case 48 To 57, 65 To 90, 97 To 122
                sRes = sRes & sChar
            Case Else
                sRes = sRes & "%" & Hex(Asc(sChar))
        End Select
    Next i
    URLEncode = sRes
End Function

