Attribute VB_Name = "ModSynchronisation"

Option Explicit
Sub Alerte2j(Optional afficheMessage = True)
    '*************************************************
    '* Cette procédure renvoie toutes les périodes
    '* de deux jours continus dans les calendriers airbnb
    '*************************************************
    ThisWorkbook.Activate
    'On met à jour les données du planificateur de la page d'accueil
    ControlePlanificateur "Alerte2j"
    
    '-----------------------------------------------------------------------------
    'Définition des variables
    '----------------------------------------------------------------------------
    Dim Logements As Variant:    Logements = Range("Logements[Logements]").value
    Dim nbLogements As Long:     nbLogements = CalculLogementsActifs
    Dim anDebut As Long:     anDebut = Year(Now)
    Dim anFin As Long:      anFin = anDebut + 1
    Dim T As Variant        'CA par jour
    T = CalculJour(anDebut, anFin)
    Dim U As Variant        'Prix par jour
    U = GetTablePrix
    Dim nbSources As Integer: nbSources = Feuil5.Range("Sources").ListObject.ListRows.Count
   
    '*************************************************************
    '* On parcourt la liste des réservations pour chaque logement
    '*************************************************************
    Dim dateDebut As Date
    Dim dateFin As Date
    dateDebut = DateSerial(anDebut, month(Now), 1)
    dateFin = DateAdd("yyyy", 1, dateDebut) - 1
    
    'On ajuste besoin de savoir si le logement est pris
    Dim i As Long
    Dim j, k As Integer
    
    'On calcule le CA en additionnant pour chaque source de revenus
    For i = dateDebut To dateFin
        For j = 1 To nbLogements
            For k = 1 To nbSources
                T(i, j, 0) = T(i, j, 0) + T(i, j, k)
            Next k
        Next j
    Next i
    
    '*On lance l'exploration
    Dim texte As String
    
    For i = LBound(U, 2) + 1 To dateFin - 2
        For j = 1 To nbLogements
            If T(i - 1, j, 0) <> 0 And T(i, j, 0) = 0 And T(i + 1, j, 0) = 0 And T(i + 2, j, 0) <> 0 Then
                'On  a les conditions de deux jours consécutifs libres
                texte = texte + Range("Logements[Logements]")(j).value + " : 2 nuits à partir du " + CStr(CDate(i)) + " (Prix airBnb = " + CStr(U(j - 1, i)) + " €)" + vbCrLf
            End If
        Next j
    Next i
    
    'On crée et on envoir un message
    
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = "christophe@tricaud.com"
        '.CC = "contact@automatiseoffice.com; tonadresse@gmail.com"
        '.BCC = "" 'champ mail en copie caché
        .Subject = "Relevé des locations sur 2 jours"
        .body = "Voici le relevé des périodes de deux jours disponibles:" & vbCrLf & vbCrLf & texte
                
        '.Display 'affiche le mail en brouillon dans Outlook, pratique
        'pour vérifier avant d'envoyer
        .send                                    'envoie directement le mail
        '.Save 'sauvegarde le mail
    
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
  
    
    'If afficheMessage Then MsgBox "Envoi du mail réalisé."
 
End Sub
Sub AppelNotificationsPulse()
    Range("LogExtraction") = ""
    NotificationsPulse
End Sub

Sub calculM()
    Sheets("Ménages").Activate
End Sub


Sub NotificationsPulse()
    '-------------------------------------------------------------
    'Permet de gérer les notifications issues de pulse
    '-------------------------------------------------------------
    Dim Notif As String
    Dim T(4)
    ThisWorkbook.Activate
    ControlePlanificateur "NotificationsPulse"
    
    
    '-------------------------------------------------------------
    '1.1 On traite Booking
    '-------------------------------------------------------------
    Notif = Trim(GetDataFromKVStore("pulse"))
    If Notif <> "" Then
        'On efface la notification pour la prochaine fois
        SendDataToKVStore "pulse", " "

        '-------------------------------------------------------------
        '1.2 On traite la notification
        '-------------------------------------------------------------
        T(1) = Format(Now, "dd-mm-yy hh:nn:ss")
        T(2) = "Booking"
        T(3) = ""
        T(4) = Notif

        MajNotifs T
    
    End If



    '-------------------------------------------------------------
    '2.1 On traite AirBnb
    '-------------------------------------------------------------
    Notif = Trim(GetDataFromKVStore("airbnb"))
    If Notif <> "" Then
        'On efface la notification pour la prochaine fois
        SendDataToKVStore "airbnb", " "

        '-------------------------------------------------------------
        '2.2 On traite la notification
        '-------------------------------------------------------------
        T(1) = Format(Now, "dd-mm-yy hh:nn:ss")
        T(2) = "Airbnb"
        T(3) = ""
        T(4) = Notif
        
        'On traite les différents cas
        
        Select Case True
            Case IsNumeric(Left(T(1), 1)): T(3) = "Message"
            Case InStr(T(1), "Annulation") = 1: T(3) = "Réservation"
            Case InStr(T(1), "a annulé sa réservation") > 0: T(3) = "Réservation"
            Case InStr(T(1), "Un versement") > 0: T(3) = "Paiement"
             Case InStr(T(1), "Laissez un commentaire") > 0: T(3) = "Notation"
            
         End Select

        MajNotifs T
    
    End If
    
    
    
    '-------------------------------------------------------------
    '3.1 On traite les mails Airbnb
    '-------------------------------------------------------------
    Dim emails As Collection
    Dim DateMail As Date
    Dim texte As String
    Dim Id As String
    Dim AccessToken As String
   
    AccessToken = GetAccessToken
   
    Set emails = New Collection
    Set emails = ProcessEmails
    Dim email As Object
    If Not emails Is Nothing Then
        For Each email In emails
            texte = email("snippet")
            If Left(UCase(texte), 4) <> "RE :" Then
                Id = email("id")
                DateMail = ConvertTimestampToDate(email("date"))
              
                Dim U As Object
                Set U = Range("Notifs").ListObject.ListRows.Add(1)
              
                U.Range(1, 2) = "Airbnb"
                U.Range(1, 1) = Format(DateMail, "dd-mm-yy hh:nn")
                U.Range(1, 4) = texte
                U.Range(1, 5) = "Email"
                U.Range(1, 6) = "X"
            End If
            
            'On supprime l'email
            DeleteEmail AccessToken, Id
            
        Next
    End If
    
    '-------------------------------------------------------------
    '9. On relance la boucle d'attente et lecture
    '-------------------------------------------------------------
    ApplicationOntime Now + TimeValue("00:10:00"), "NotificationsPulse"
    Feuil19.Range("execNotif") = Format(Now, "dd-mm-yy hh:nn:ss")
End Sub
Sub MajNotifs(T As Variant)
    '-----------------------------------------------------------
    'Cette procédure ajoute une notification au tableau
    '-----------------------------------------------------------
    RAZFiltres ("Notifs")
    
    Dim U As Object
    Dim v As Variant
    Set U = Feuil19.Range("Notifs").ListObject.ListRows.Add(1)
    v = Split(T(4), " | ")
      
    U.Range(1, 2) = T(2)
    U.Range(1, 1) = v(0) + " " + v(1)
    U.Range(1, 4) = v(2) + vbCrLf + Left(v(3), 75)
    U.Range(1, 5) = "App"
    U.Range(1, 6) = "X"

End Sub

Sub PlanificationGlobale(Optional Repetition = True)
'---------------------------------------------------
'Cette procédure permet de mettre à jour toutes les donées
'---------------------------------------------------
    ThisWorkbook.Activate
    
    ControlePlanificateur "PlanificationGlobale"
    Range("LogExtraction") = ""
    
    '---------------------------------------------------
    '1. Mise à jour du fichier
    '---------------------------------------------------
    Dim Titre As String
    Titre = "PLANIFICATION DU " + Format(Now, "dd-mm-yyyy hh:nn:ss")
    log Titre
    
    'Log "-------------------------------------"
    'Log "A. SAUVEGARDE DU FICHIER"
    'Log "-------------------------------------"
   Sauvegarde
    'Log ""
    
    log "-------------------------------------"
    log "A. TRAITEMENTS AIRBNB"
     log "-------------------------------------"
   AirBnbGetReservations False
    log ""
    
      log "-------------------------------------"
    log "B. TRAITEMENTS BOOKING"
     log "-------------------------------------"
   BookingGetReservations False
    log ""
    
    log "-------------------------------------"
    log "C. TRAITEMENTS GUESTY"
     log "-------------------------------------"
   GuestyGetReservationsOnline False
    log ""
    
  'Log "-------------------------------------"
   ' Log "E. MISE A JOUR CALCUL FEUILLE"
   ' Log "-------------------------------------"
    Feuil11.Activate
    log ""
    LogCA12Mois
    
     'Log "-------------------------------------"
    'Log "F. ENVOI MAIL - FIN DE PLANIFICATION"
    'Log "-------------------------------------"
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = "christophe.tricaud@gmail.com"
        .Subject = Titre
        .body = Range("LogExtraction")
        .send
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
    Feuil8.Activate
    log ""
   
   If Repetition Then ApplicationOntime TimeValue("07:00:00"), "PlanificationGlobale"
End Sub
Sub Alerte2JB()
    'Procédure appelée par batch le lundi matin
    '1. On met à jour les réservations airbnb
    ModAirbnb.AirBnbGetReservations

    '2. On lmance le mail
    Alerte2j False
    
    '3. On ferme le classeur
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Application.Quit
End Sub

Sub Programmation()
'---------------------------------------------
'Permet de lancer la programmation
'---------------------------------------------
    Dim Repetition As Boolean
    
    Repetition = MsgBox("Souhaitez vous lancer la programmation quotidienne ?", vbYesNo)
   If Repetition = vbNo Then
       PlanificationGlobale False
    Else
        'On met la planification en place
        ApplicationOntime TimeValue("06:30:00"), "PlanificationGlobale"
    End If

    
End Sub

