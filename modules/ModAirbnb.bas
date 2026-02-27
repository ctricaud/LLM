Attribute VB_Name = "ModAirbnb"
Option Explicit
Dim AirbnbExisteErreur
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const VK_CONTROL As Byte = &H11
Const VK_SUBTRACT As Byte = &H6D ' Code pour la touche "-"
Const KEYEVENTF_KEYUP As Long = &H2
Const VK_PGDN As Byte = &H22 ' Code pour la touche Page Down
Const VK_PGUP As Byte = &H21 ' Code pour la touche Page Up
Function AirBnbGetCode(lastKey As String) As String
    Dim code As String
    Do
        code = ExtractAirbnbCode(GetValueFromBin(JSONBinAirBNB, "Airbnb"))
        Application.Wait (Now + TimeValue("0:00:05"))
        DoEvents
    Loop Until code <> lastKey
    AirBnbGetCode = code
End Function

Sub airbnbGetStats(Optional driver As Object)
    '--------------------------------------------------------------
    'Cette procédure permet de récupérer toutes les stats
    'Pour les logements sur airbnb
    '--------------------------------------------------------------
    '1. On définit les informations de connexions sur le driver
    '--------------------------------------------------------------
    Const idConversion = 1
    Const idDelai = 2
    Const idVue = 3
    Const idWishList = 4
    
    Dim T() As Variant
    
    Dim url(4)
    Dim urlAppend0, urlAppend As String
    url(idConversion) = "https://www.airbnb.fr/performance/conversion/conversion_rate"
    url(idDelai) = "https://www.airbnb.fr/performance/conversion/booking_window"
    url(idVue) = "https://www.airbnb.fr/performance/conversion/p3_impressions"
    url(idWishList) = "https://www.airbnb.fr/performance/conversion/wishlist"
    
    urlAppend0 = "?lid%5B%5D=&ds-start=&ds-end="
    
     Dim jourDebut As Integer
     Dim lastJour
     'On recherche le dernier enregistrement
     lastJour = Range("statsAirbnbApollinaire[date]")(Range("statsAirbnbApollinaire[date]").Count)
     jourDebut = DateDiff("d", Date, lastJour) + 1
     If jourDebut > 0 Then Exit Sub
     

    
     '--------------------------------------------------------------
    '2. On récupère les informations sur les logements
    ' qui nous intéressent
    '--------------------------------------------------------------
    Dim Logements As Variant
    Dim nbLogements As Integer
    Logements = Range("Logements")
    nbLogements = CalculLogementsActifs
    nbLogements = 2
    Dim reservations As Variant
    Dim iReservation, nbReservations, diff As Long
    Dim idBookingDate As Integer
    idBookingDate = Range("ListeRésas").ListObject.ListColumns("booking_Date").Index
    reservations = Range("ListeRésas")
    
     '--------------------------------------------------------------
    '3. On lance le driver s'il n'existe pas
    '--------------------------------------------------------------
    If driver Is Nothing Then
        Set driver = AirbnbLaunchEdge
    End If
    
    '--------------------------------------------------------------
    '4. On boucle sur chaque logements
    '--------------------------------------------------------------
    Dim iLogement, iJour As Integer
    Dim NomTableau As String
    Dim elems As Variant
    
    For iLogement = 1 To nbLogements
        NomTableau = "StatsAirbnb" + Logements(iLogement, 1)
            
        For iJour = jourDebut To 0
            '--------------------------------------------------------------
           ' 4.1 On prépare la ligne à insérer
            '--------------------------------------------------------------
            ReDim T(1 To Range(NomTableau).Columns.Count) As Variant
            
            'La date de calcul
            T(Range(NomTableau).ListObject.ListColumns("Date").Index) = DateAdd("d", iJour, Date)
            urlAppend = Replace(urlAppend0, "?lid%5B%5D=", "?lid%5B%5D=" + Logements(iLogement, 2))
            urlAppend = Replace(urlAppend, "&ds-start=", "&ds-start=" + CStr(iJour - 30))
            urlAppend = Replace(urlAppend, "&ds-end=", "&ds-end=" + CStr(iJour))
            
            '--------------------------------------------------------------
           ' 4.2 Récupération des taux de conversion
            '--------------------------------------------------------------
            driver.get url(idConversion) + urlAppend
            driver.Wait 2000
             
            Set elems = driver.FindElementsByClass("_1wp8t9")
            
             T(Range(NomTableau).ListObject.ListColumns("Conversion").Index) = elems(1).Text
             T(Range(NomTableau).ListObject.ListColumns("Premiere page").Index) = elems(2).Text
             T(Range(NomTableau).ListObject.ListColumns("%1").Index) = elems(3).Text
             T(Range(NomTableau).ListObject.ListColumns("%2").Index) = elems(4).Text
             
           '--------------------------------------------------------------
           ' 4.3 récupération des vues
            '--------------------------------------------------------------
             driver.get url(idVue) + urlAppend
             driver.Wait 2000
             Set elems = driver.FindElementsByClass("_1wp8t9")
             
              T(Range(NomTableau).ListObject.ListColumns("Vues").Index) = elems(1).Text
              T(Range(NomTableau).ListObject.ListColumns("Impressions").Index) = elems(2).Text
              
              
           '--------------------------------------------------------------
           ' 4.4 mise en favori
            '--------------------------------------------------------------
             driver.get url(idWishList) + urlAppend
             driver.Wait 2000
             Set elems = driver.FindElementsByClass("_1wp8t9")
             
              T(Range(NomTableau).ListObject.ListColumns("Favoris").Index) = elems(1).Text
              
              
           '--------------------------------------------------------------
           ' 4.5 Delai
            '--------------------------------------------------------------
             driver.get url(idDelai) + urlAppend
             driver.Wait 3000
             Set elems = driver.FindElementsByClass("_1wp8t9")
             
              T(Range(NomTableau).ListObject.ListColumns("Delai").Index) = Left(elems(1).Text, 3)
               
               
           '--------------------------------------------------------------
           ' 4.6 réservation
            '--------------------------------------------------------------
             nbReservations = 0
             
             For iReservation = 1 To UBound(reservations)
                diff = DateDiff("d", iJour, reservations(iReservation, idBookingDate))
                If diff > -31 And diff < 1 And Logements(iLogement, 1) = reservations(iReservation, 1) Then
                    nbReservations = nbReservations + 1
                End If
                If diff < -60 Then Exit For
             Next iReservation
             
              T(Range(NomTableau).ListObject.ListColumns("Reservations").Index) = nbReservations
               
              '--------------------------------------------------------------
              ' 4.9 On ajoute au tableau dont on a besoin
             '--------------------------------------------------------------
            'On vérifie si le calcul a déjà été fait ou pas
            If T(Range(NomTableau).ListObject.ListColumns("Date").Index) <> Range(NomTableau + "[Date]")(Range(NomTableau + "[Date]").Count) Then
                Range(NomTableau).ListObject.ListRows.Add
            End If
            Range(NomTableau).Rows(Range(NomTableau).Rows.Count).value = T
            
        Next iJour
    Next iLogement
    
    
End Sub

Function ExtractAirbnbCode(msg As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")  ' late-binding
    
    With re
        .Pattern = "\b\d{6}\b"   ' six chiffres isolés
        .Global = False          ' on ne veut que le premier match
        .IgnoreCase = True
    End With
    
    If re.Test(msg) Then
        ExtractAirbnbCode = re.Execute(msg)(0).value
    Else
        ExtractAirbnbCode = ""   ' pas de code trouvé
    End If
End Function
Function AirbnbLaunchEdge() As WebDriver
    Dim driver As New WebDriver
      
    ' Connexion à airbnb hosting
    driver.Start "chrome"
    driver.Window.Maximize
    driver.get "https://www.airbnb.fr/hosting"
    'driver.ExecuteScript "document.cookie = '_aat=" + Range("Cookie") + "';"
    'driver.get "https://www.airbnb.fr/hosting"
    driver.Wait 2000
    
    'On récupère l'anciienne clé pour vérifier quand elle aura changé
    Dim lastKey As String
    lastKey = ExtractAirbnbCode(GetValueFromBin(JSONBinAirBNB, "Airbnb"))
    
    'On rentre le numéro de téléphone pour récupérer le code
    Dim elem As Object
    Set elem = driver.FindElementById("phoneInputphone-login")
    elem.Clear
    elem.SendKeys "658286395"
    driver.ExecuteScript "arguments[0].click();", driver.FindElementByXPath("//span[contains(@class,'t1dqvypu')]")

    driver.Wait 2000
    'On enlève le rgpd si nécessaire
    On Error Resume Next
        driver.FindElementByXPath("//button[normalize-space()='Accepter tout']").Click
    On Error GoTo 0
    
    
    'On récupère le code envoyé et in le saisit
    Dim code As String
    code = AirBnbGetCode(lastKey)
    
    Set elem = driver.FindElementById("phone-verification-code-form__code-input")
    elem.Clear
    elem.SendKeys code
    
    'On enlève le rgpd si nécessaire
    On Error Resume Next
        driver.FindElementByXPath("//button[normalize-space()='Accepter tout']").Click
    On Error GoTo 0
    
    Set AirbnbLaunchEdge = driver
End Function
Sub AirBnbAcces()
    '----------------------------------------------------
    'Cette procédure permet de mettre à jour le cookie
    'd'accès à AirBNB si il n'est plus valable
    '----------------------------------------------------
    '1. On vérifie si l'accès est possible ou pas
    '----------------------------------------------------
    'Log "1. Vérification de l'accès AirBnb"
    'Log "------------------------"
    Dim xmlhttp As New MSXML2.ServerXMLHTTP60
    Dim myUrlBase As String
    myUrlBase = "https://www.airbnb.fr/api/v2/reservations?locale=fr&currency=EUR&_format=for_remy&_limit=40&collection_strategy=for_reservations_list&sort_field=start_date&sort_order=desc&status=accepted%2Crequest%2Ccanceled"
    myUrlBase = myUrlBase + "&key=" + Range("Key")
     Dim retour As String
    xmlhttp.Open "GET", myUrlBase, False
    xmlhttp.setRequestHeader "Cookie", "_aat=" + Range("Cookie")
    xmlhttp.send
                    
    retour = xmlhttp.responseText
    If InStr(retour, """authentication_required""") = 0 Then
        Exit Sub
    End If
    
    '----------------------------------------------------
    '2. On va mettre à jour le cookie qui n'est plus valable
    '----------------------------------------------------
    Dim driver As New WebDriver

    driver.Start "chrome"
    driver.get "https://www.airbnb.fr/hosting"
    driver.Window.Maximize
    WaitReadyState driver
    driver.Wait 2000

    '------------------------------------------------------------------------
    '3. On supprime le RGPD
    '------------------------------------------------------------------------
    driver.FindElementByXPath("//button[normalize-space()='Accepter tout']").Click

    '------------------------------------------------------------------------
    '4. On récupère le dernier code stocké
    '------------------------------------------------------------------------
    Dim lastKey As String
    lastKey = GetValueFromBin(JSONBinAirBNB, "Airbnb")
    
    '------------------------------------------------------------------------
    '5. On rentre le numéro de téléphone
    '------------------------------------------------------------------------
    driver.FindElementById("phoneInputphone-login").SendKeys "658286395"
    driver.FindElementByXPath("//button[span[contains(text(), 'Continuer')]]").Click

    driver.Wait 5000
    
     '------------------------------------------------------------------------
    '6. On récupère le code dans jsonBin
    '------------------------------------------------------------------------
   Dim code As String
   code = AirBnbGetCode(lastKey)
   
    '------------------------------------------------------------------------
    '7. On remplit le champs et on valide
    '------------------------------------------------------------------------
    driver.FindElementById("phone-verification-code-form__code-input").SendKeys code
    On Error Resume Next
    driver.FindElementByXPath("//button[contains(text(), 'Continuer')]").Click
    On Error GoTo 0
    
    '------------------------------------------------------------------------
    '8. On récupère la valeur du cookie
    '------------------------------------------------------------------------
    Dim cookieValue
    Dim cookie
    
    For Each cookie In driver.Manage.Cookies
         If cookie.Name = "_aat" Then
            cookieValue = cookie.value
            Exit For
        End If
    Next cookie
    
    Range("Cookie") = cookieValue
End Sub
