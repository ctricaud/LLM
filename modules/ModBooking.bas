Attribute VB_Name = "ModBooking"

'Option Explicit
Sub BookingGetAccess(driver As WebDriver)
    '------------------------------------------------------------------------
    'Permet de rentrer les logins et les mots de passe pour initier la connexion
    '------------------------------------------------------------------------
    driver.get "https://admin.booking.com"
    WaitReadyState driver
    driver.Wait 2000
    
    Dim Script As String
    Script = "Object.defineProperty(navigator, 'webdriver', {get: () #  undefined});"

    ' Injecter le script JavaScript
    'Driver.ExecuteScript Script

    ' Simuler une pause pour éviter la détection d'automatisation (comportement humain)
    'Application.Wait Now + TimeValue("00:00:02")

    
    '------------------------------------------------------------------------
    '1. On supprime le RGPD
    '------------------------------------------------------------------------
    driver.FindElementById("onetrust-accept-btn-handler").Click
    driver.Wait 2000
    
    '------------------------------------------------------------------------
    '2. On rempli les crédits pour se connecter
    '------------------------------------------------------------------------
    'Le login
    Dim InputField As WebElement
    Set InputField = driver.FindElementById("loginname")
    Dim Login As String: Login = "ac@tricaud.com"
    InputField.SendKeys Login
    driver.Wait 2000
    
    'Le mot de passe
    driver.FindElementByXPath("//button[.//span[text()='Next']]").Click
    WaitReadyState driver
    driver.Wait 2000
    
    On Error Resume Next
    Do
        Err = 0
        Set InputField = driver.FindElementById("password")
    Loop Until Err = 0
    On Error GoTo 0
    
    InputField.SendKeys "Corne12ct@ctct"
    driver.FindElementByXPath("//button[.//span[text()='Sign in']]").Click

    WaitReadyState driver
    driver.Wait 5000
    On Error Resume Next
        driver.FindElementById("onetrust-accept-btn-handler").Click
    On Error GoTo 0
    


 

End Sub


Sub BookingCalculCommissionMensuelle()
    '***********************************************************
    '* Cette procédure envoie un mail au concierge pour
    '* Restituer le montant de la commission qui est à facturer
    '***********************************************************
    '* 1. On sélectionne le mois qui nous intéresse
    '***********************************************************
    Load frmEmailCommissionBookingHobe
    frmEmailCommissionBookingHobe.Show
End Sub

Sub BookingTraitementFichierFactures(Optional effaceLog = True)
    '-----------------------------------------------
    'On traite les factures récupérées sur le site de Booking
    '-----------------------------------------------
    If effaceLog Then Range("LogExtraction") = ""
    
    '----------------------------------------------------
    '1. Définition et mise en place
    '----------------------------------------------------
    Dim LignesFichiers(): ReDim LignesFichiers(20000)
    Dim ListeResas As Variant: ListeResas = Range("ListeRésas").ListObject.DataBodyRange.value
    Dim iLigne As Long
    
    ' Récupérer le chemin du répertoire depuis la cellule "Download"
    Dim cheminDossier As String: cheminDossier = Range("DirDownload").value
    If Right(cheminDossier, 1) <> "\" Then cheminDossier = cheminDossier & "\"
    Dim fso As Object
    Dim MonFichier As String
    Dim IndexFichier As Integer
    Dim nbLignes As Integer
    Dim fileItem
    Dim Titre As String
    
    Dim fichier As String: fichier = dir(cheminDossier & "*statements*.csv")
    
    '------------------------------------------------------------
    '2. On lance la boucle de lecture des différents fichiers
    '------------------------------------------------------------
    Do While fichier <> ""
        '* On récupère le fichier et on le met dans une variable
        Set fso = CreateObject("Scripting.FileSystemObject")
        MonFichier = cheminDossier & fichier
    
        'Set fileItem = fso.GetFile(MonFichier)
        'Log ("Factures booking du : " + Format(fileItem.DateLastModified, "dd/mm/yyyy hh:mm:ss"))
        'Set fileItem = Nothing
        'Set fso = Nothing
   
        IndexFichier = FreeFile()
        Open MonFichier For Input As #IndexFichier
 
        While Not EOF(IndexFichier)
            nbLignes = nbLignes + 1
            Line Input #IndexFichier, LignesFichiers(nbLignes)
            LignesFichiers(iLigne) = Replace(LignesFichiers(nbLignes), Chr(172), "€")
        Wend
 
        Close #IndexFichier
        
        'On supprime le fichier traité
        
        'On passe au suivant
        fichier = dir
    Loop
    
    ReDim Preserve LignesFichiers(nbLignes)
    Dim nbFactures As Long
    Dim tableLigne As Variant
    Dim numero As Variant
    Dim reversement As Variant
    Dim i, j, trouve As Long
    Dim dateDebut, dateFin As Date
    Dim nbNuits As Integer
    
        For iLigne = 2 To UBound(LignesFichiers)
            LignesFichiers(iLigne) = ConvertirEncodage(ByVal LignesFichiers(iLigne))
           
            If InStr(LignesFichiers(iLigne), "Réservation") = 1 Or InStr(LignesFichiers(iLigne), "Reservation") = 1 Then
                nbFactures = nbFactures + 1
                tableLigne = Split(LignesFichiers(iLigne), ",")
                
                'On recherche la réservation dans la liste des réservations
                If tableLigne(2) = "" Then
                    log "Un problème est survenu avec la réservation " + tableLigne(1) + vbCrLf _
                        + "Montant : " + CStr(tableLigne(12))
                Else
                    dateDebut = CDate(MoisUSDate(Replace(tableLigne(2), Chr(34), "")))
                    dateFin = CDate(MoisUSDate(Replace(tableLigne(3), Chr(34), "")))
                    nbNuits = CLng(dateFin) - CLng(dateDebut)
                    reversement = Replace(tableLigne(12), ".", ",")
                     trouve = False
                    For j = 1 To UBound(ListeResas)
                        If ListeResas(j, 4) = nbNuits And ListeResas(j, 2) = "Booking" And Int(ListeResas(j, 3)) = dateDebut Then
                            trouve = True
                            'On regarde si le paiement est validé
                            If ListeResas(j, 12) = "" Then
                                'Sinon on vérifie le montant
                                If Abs(reversement - ListeResas(j, 10) - ListeResas(j, 9)) < 5 Then
                                    'On valide le montant
                                    ListeResas(j, 12) = "ü"
                                    Titre = "# Validation du paiement de la réservation " & CStr(numero) & _
                                        " pour " & ListeResas(j, 1) & " du " & Format(ListeResas(j, 3), "dd-mm-yyyy") & _
                                        " pour un montant de " & CStr(reversement) & " €" + vbCrLf
                                        
                                    log Titre
                                    Notification Date, "Booking", Titre, "Paiement", "Automate"
                        
                                Else
                                    'On affiche une alerte
                                    Titre = "# La facture suivante n'a pas été affectée : " & CStr(numero) & " - " & CStr(reversement)
                                    Titre = Titre + vbCrLf + "Le montant ne correspond pas." + vbCrLf
                                    
                                    log Titre
                                    Notification Date, "Booking", Titre, "Réservation", "Automate"
                
                                End If
                            End If
                        End If
                    
                    Next j
                
                    If Not trouve Then
                        Titre = "# La facture suivante n'a pas été affectée : " & CStr(numero) & " - " & CStr(reversement) + vbCrLf
                        Titre = Titre + " du " + Format(dateDebut, "dd/mm/yyyy") + " au " + Format(dateFin, "dd/mm/yyyy") + vbCrLf
                        
                        log Titre
                        Notification Date, "Booking", Titre, "Réservation", "Automate"
                
                    End If
                End If
            End If
        Next iLigne

    'Log CStr(nbFactures) + " Factures récupérées"
    Range("ListeRésas").ListObject.DataBodyRange.value = ListeResas
    SupprimerFichiers "*statement*.csv"
End Sub
Sub BookingGetInvoices(Optional driver As Object)
    '------------------------------------------------------------------------
    'Récupération des factures en téléchargeant les fichiers nécessaires
    '------------------------------------------------------------------------
    '1. On récupère les logements concernés id Booking_id
    '------------------------------------------------------------------------
    Dim Traitement As Boolean
    
    If driver Is Nothing Then
        Traitement = True
        logExtraction = ""
        Set driver = New WebDriver
        driver.SetCapability "chromeOptions", _
                             "{'args': ['--disable-blink-features=AutomationControlled', " & _
                             "'--disable-infobars', '--no-sandbox', '--disable-dev-shm-usage', '--start-maximized']}"
        driver.SetCapability "excludeSwitches", Array("enable-automation")
        driver.SetCapability "useAutomationExtension", False
            
        driver.Start "chrome"
        BookingGetAccess driver
    End If
    
    Dim Logements As Variant: Logements = Range("Logements").value
    Dim iLogement As Integer
    Dim URLBase As String
    Dim url As String
    Dim ses As String
    Dim iYear As Integer
    
    Dim dropdownButton As Variant
    Dim Options As Variant
    Dim optionElement
    
    URLBase = "https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/payouts.html?lang=fr&"
    ses = BookingExtractSes(driver)
    log "Téléchargement des factures de booking"
    
    For iLogement = 1 To UBound(Logements)
        If Logements(iLogement, 6) <> "" Then
            '------------------------------------------------------------------------
            '2. On télécharge les factures du logement
            '------------------------------------------------------------------------
            url = URLBase + "ses=" + ses + "&hotel_id=" + CStr(Logements(iLogement, 6))
            driver.get url
            WaitReadyState driver
            driver.Wait 1000
            On Error Resume Next
                driver.FindElementById("onetrust-accept-btn-handler").Click
         On Error GoTo 0
        
        driver.FindElementByXPath("//span[text()='Télécharger tous les rapports']").Click
        driver.Wait 5000
            
    End If
    Next iLogement
    
    If Traitement Then BookingTraitementFichierFactures
    
End Sub

Function BookingExtractSes(driver As WebDriver) As String
    Dim url As String: url = driver.url
    
   url = Mid(url, InStr(url, "ses=") + 4)
   If InStr(url, "&") = 0 Then
        BookingExtractSes = url
   Else
        BookingExtractSes = Left(url, InStr(url, "&") - 1)
   End If
End Function


