Attribute VB_Name = "ModGuesty"
Option Explicit
Sub GuestyGetPrixOnline()
    '------------------------------------------
    'Récupération des prix sur Guesty
    '------------------------------------------
    'Range("LogExtraction") = ""
    'Log "---------------------------------------"
    'Log "Récupération des prix Guesty"
    'Log "---------------------------------------"
    
    '------------------------------------------
    '0. Chargement des paramètres
    '------------------------------------------
    Dim httpRequest As Object
    Dim url As String
    Dim responseText As String
    Dim dateDebut As String
    Dim dateFin As String
    
    Dim Logements: Logements = Range("Logements").value
    Dim nbLogement: nbLogement = UBound(Logements)
    Dim iLogement As Integer
    Dim T() As Variant
    ReDim T(CLng(Date) To CLng(DateAdd("d", 365, Date)), nbLogement)
    
    Dim dateJour As Variant
    Dim statusJour As String
    Dim priceJour As Variant
    Dim iJour As Integer
    
    '---------------------------------------
    '1. On récupère le JSON sur guesty
    '---------------------------------------
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    For iLogement = 1 To nbLogement
        If Logements(iLogement, 7) <> "" Then
            'Log "Récupération des prix pour " + Logements(iLogement, 1)
        
            '1.1 Définir l'URL et la clé API
            url = Range("GuestyURLPrix")
            '1.1.a On met à jour l'id du logement
            url = Replace(url, "idLogement", Logements(iLogement, 7))
            '1.1.2 On met à jour les dates d'extraction
            url = Replace(url, "dateFin", ConvertirDate(DateAdd("d", 365, Date), "-"))
            url = Replace(url, "dateDebut", ConvertirDate(Date, "-"))
    
            '1.2 Ouvrez la requête HTTP GET
            httpRequest.Open "GET", url, False
            
            '1.3 Définir l'en-tête requis
            httpRequest.setRequestHeader "Authorization", Range("GuestyKey")
    
            '1.4 Envoyez la requête
            httpRequest.send
    
            '1.5 Vérifiez le statut de la réponse
            If httpRequest.Status <> 200 Then
                GuestyGetAccessToken
                GuestyGetPrixOnline
                Exit Sub
            End If
    
            '1.6 Réponse réussie
            responseText = httpRequest.responseText
            Dim dictionnaire As Variant
            Set dictionnaire = ParseJson(responseText)
        
            '------------------------------------------
            '2. On récupère les prix dans la table T
            '------------------------------------------
            For iJour = 0 To 365
                dateJour = dictionnaire("obj(" + CStr(iJour) + ").date")
                dateJour = CDate(ConvertirDate(dateJour))
                statusJour = dictionnaire("obj(" + CStr(iJour) + ").status")
                priceJour = dictionnaire("obj(" + CStr(iJour) + ").price")
                priceJour = CCur(priceJour)
            
                If statusJour = "available" Then
                    T(CLng(dateJour), iLogement) = priceJour
                End If
            Next iJour
            Set dictionnaire = Nothing
        End If
    Next iLogement
    
    '------------------------------------------
    '2. On met à jour le tableau
    '------------------------------------------
    Dim i As Long
    For i = LBound(T) To UBound(T): T(i, 0) = i: Next i
    TableToTableau T, "TableauPrix"
  
    '9. Libérez l'objet
    Set httpRequest = Nothing
    
End Sub

Sub GuestyGetAccessToken()
    ' ------------------------------------
        ' ------------------------------------
        '2. On envoie la demande de token à Guesty
        '------------------------------------
        Dim http As Object
        Dim url As String
        Dim ClientId As String
        Dim ClientSecret As String
        Dim HostName As String
        Dim APIKey
        
        ' 2.1 Initialiser les variables
        '------------------------------------
        ClientId = "christophe.tricaud@gmail.com"
        ClientSecret = "Corne12ct@"
        url = "https://app.guesty.com/api/owners/auth/login"
        HostName = "owner.wehobe.com"
        APIKey = "KdDnXIZcVp0HBF2pzzVDHvNqE0CwxVnp"
    
        ' 2.2 Créer un objet HTTP et envoie de la requête
        '------------------------------------
        Set http = CreateObject("MSXML2.XMLHTTP")
    
        ' Préparer la requête HTTP POST
        http.Open "POST", url, False
        http.setRequestHeader "Accept", "application/json"
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
        ' Construire le corps de la requête avec les paramètres codés en URL
        Dim postData As String
        postData = "apiKey=" & APIKey & _
                   "&hostname=" & HostName & _
                   "&username=" & ClientId & _
                   "&password=" & ClientSecret
    
        ' Envoyer la requête avec les données
        http.send postData
    
        ' 2.3 Vérifier et traitement de la réponse HTTP
        '------------------------------------
        If http.Status = 200 Then
            Dim JsonResponse As Object
            Dim Response As String
            
            Response = http.responseText
            
            ' 2.4 traiter la réponse JSON
             '------------------------------------
            Set JsonResponse = ParseJson(Response)
        
            ' Extraire le token d'accès (access_token)
            Dim AccessToken As String
            AccessToken = "Bearer " + JsonResponse("obj.token")
            
            Range("GuestyKey") = AccessToken
        Else
            '------------------------------------
            '3. Erreur de connexion
            MsgBox "Erreur: " & http.Status & " - " & http.responseText
        End If
    
        ' Libérer l'objet HTTP
        Set http = Nothing
 

End Sub






