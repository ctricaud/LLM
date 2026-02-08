Attribute VB_Name = "ModKVStore"

Option Explicit


Sub CreateCollection()
    Dim httpRequest As Object
    Dim url As String
    Dim APIKey As String
    Dim postData As String
    Dim responseText As String
    
    ' Initialisez l'objet HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Définissez l'URL et la clé API
    url = "https://api.kvstore.io/collections"
    APIKey = Range("KVStoreKey")
    
    ' Données à envoyer en format JSON
    postData = "{""collection"" : ""CTCOLLECTION""}"
    
    ' Ouvrez la requête HTTP POST
    httpRequest.Open "POST", url, False
    
    ' Définissez les en-têtes requis
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "kvstoreio_api_key", APIKey
    
    ' Envoyez la requête avec les données JSON
    httpRequest.send postData
    
    ' Récupérez la réponse du serveur
    responseText = httpRequest.responseText
    
    ' Affichez la réponse (facultatif)
    MsgBox "Réponse du serveur : " & responseText
    
    ' Nettoyez l'objet
    Set httpRequest = Nothing
End Sub


Function GetDataFromKVStore(Optional cle = "sms")
    Dim httpRequest As Object
    Dim url As String
    Dim APIKey As String
    Dim responseText As String
    
    ' Initialisez l'objet HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Définissez l'URL et la clé API
    url = "https://api.kvstore.io/collections/CTCOLLECTION/items/cle_" + cle
    APIKey = Feuil5.Range("KVStoreKey")
    
    ' Ouvrez la requête HTTP GET
    httpRequest.Open "GET", url, False
    
    ' Définissez l'en-tête requis
    httpRequest.setRequestHeader "kvstoreio_api_key", APIKey
    
    ' Envoyez la requête
    httpRequest.send
    
    ' Vérifiez le statut de la réponse
    If httpRequest.Status = 200 Then
        ' Réponse réussie
        responseText = httpRequest.responseText
        Dim dictionnaire As Variant
        Set dictionnaire = ParseJson(responseText)
        
        GetDataFromKVStore = dictionnaire("obj.value")
     
    Else
        ' Gestion des erreurs
        log Format(Now, "dd-mm hh:nn:ss") + " Erreur " & httpRequest.Status & ": " & httpRequest.statusText
    End If
    
    ' Libérez l'objet
    Set httpRequest = Nothing
End Function


Sub SendDataToKVStore(Optional cle = "sms", Optional postData = "")
    Dim httpRequest As Object
    Dim url As String
    Dim APIKey As String
    Dim responseText As String
    
    ' Initialisez l'objet HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Définissez l'URL et la clé API
    url = "https://api.kvstore.io/collections/CTCOLLECTION/items/cle_" + cle
    APIKey = Feuil5.Range("KVStoreKey")
    
    ' Données à envoyer
   
    
    ' Ouvrez la requête HTTP PUT
    httpRequest.Open "PUT", url, False
    
    ' Définissez les en-têtes requis
    httpRequest.setRequestHeader "kvstoreio_api_key", APIKey
    httpRequest.setRequestHeader "Content-Type", "text/plain"
    
    ' Envoyez la requête avec les données
    httpRequest.send postData
    
    ' Récupérez la réponse du serveur
    responseText = httpRequest.responseText
    
    ' Affichez la réponse (facultatif)
    'MsgBox "Réponse du serveur : " & responseText
    
    ' Nettoyez l'objet
    Set httpRequest = Nothing
End Sub

