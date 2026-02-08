Attribute VB_Name = "ModJSONBin"

Option Explicit

' Module JSONBinExtended.bas
' Fonctions pour interagir avec JSONBin.io :
' - CreateCollection : Créer une collection
' - CreateBinInCollection : Créer un bin dans une collection
' - GetValueFromBin : Lire un champ spécifique depuis un bin
' - SetValueToBin : Mettre à jour un champ spécifique dans un bin

' Remplacez par votre propre API Key JSONBin.io
Private Const JSONBIN_API_KEY As String = "$2a$10$v3HyKEEV4MzwAK4CADMpBeP74WAV3as4bpZoid/U0.LO38eESy9yy"
Public Const JSONBinAirBNB = "6842df488a456b7966aa29b3"



' ===================================
' II. Création de bins dans une collection
' ===================================

' Crée un bin (enregistrement JSON) dans la collection spécifiée et renvoie l'ID du bin (ou "" si erreur)
' JSONContent doit être une chaîne JSON valide, ex. {"sms":"..."}
' BinName peut être vide. isPrivate = True ou False.
Public Function CreateBinInCollection( _
    ByVal CollectionId As String, _
    ByVal JSONContent As String, _
    Optional ByVal BinName As String = "", _
    Optional ByVal isPrivate As Boolean = True _
) As String
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim parsed As Object

    url = "https://api.jsonbin.io/v3/b"
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    With http
        .Open "POST", url, False
        .setRequestHeader "X-Master-Key", JSONBIN_API_KEY
        .setRequestHeader "Content-Type", "application/json"
        If Len(Trim(BinName)) > 0 Then
            .setRequestHeader "X-Bin-Name", BinName
        End If
        If Not isPrivate Then
            .setRequestHeader "X-Bin-Private", "false"
        End If
        If Len(Trim(CollectionId)) > 0 Then
            .setRequestHeader "X-Collection-Id", CollectionId
        End If
        .send JSONContent
        If .Status = 200 Or .Status = 201 Then
            responseText = .responseText
            
            Set parsed = ParseJson(responseText)
            If Not parsed Is Nothing Then
                ' L'ID du bin se trouve dans parsed("metadata")("id")
                CreateBinInCollection = parsed("metadata")("id")
            Else
                CreateBinInCollection = ""
            End If
        Else
            CreateBinInCollection = ""
        End If
    End With
End Function

' ===================================================
' III. Lecture et �criture d'un champ sp�cifique dans un bin
' ===================================================

' Lit la valeur du champ Field (ex. "sms") depuis le JSON stocké dans le bin spécifié.
' BinId : l'ID du bin (string)
' Field : la clé JSON à lire, renvoie la valeur (string) ou "" si erreur
Public Function GetValueFromBin( _
    ByVal BinId As String, _
    ByVal Field As String _
) As String
    Dim http As Object
    Dim url As String
    Dim responseText As String
    Dim parsed As Object
    Dim recordNode As Object

    url = "https://api.jsonbin.io/v3/b/" & BinId & "/latest"
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    With http
        .Open "GET", url, False
        .setRequestHeader "X-Master-Key", JSONBIN_API_KEY
        .setRequestHeader "Accept", "application/json"
        .send
        If .Status = 200 Then
            responseText = .responseText
            Set parsed = ParseJson(responseText)
            If Not parsed Is Nothing Then
                ' parsed("record") contient tout le JSON du bin
                  GetValueFromBin = parsed("obj.record.sms(0)." + Field)
            Else
                GetValueFromBin = ""
            End If
        Else
            GetValueFromBin = ""
        End If
    End With
End Function

' Met à jour la valeur du champ Field dans le bin spécifié.
' Ex. SetValueToBin("binId", "sms", "nouveau texte")
' Retourne True si succès, False sinon.
Public Function SetValueToBin( _
    ByVal BinId As String, _
    ByVal Field As String, _
    ByVal value As String _
) As Boolean
    Dim http As Object
    Dim url As String
    Dim currentJson As String
    Dim parsed As Object
    Dim recordNode As Object
    Dim updatedJson As String

    ' 1. Récupérer le JSON existant
    url = "https://api.jsonbin.io/v3/b/" & BinId & "/latest"
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    With http
        .Open "GET", url, False
        .setRequestHeader "X-Master-Key", JSONBIN_API_KEY
        .setRequestHeader "Accept", "application/json"
        .send
        If .Status <> 200 Then
            SetValueToBin = False
            Exit Function
        End If
        currentJson = .responseText
    End With

    ' 2. Parser le JSON, modifier le champ Field
    Set parsed = ParseJson(currentJson)
    If parsed Is Nothing Then
        SetValueToBin = False
        Exit Function
    End If
    ' On accède au nœud "record"
    parsed("obj.record.sms(0)." + Field) = value

    ' 3. Reconstruire le JSON à partir de recordNode
    updatedJson = "{""sms"": [{""" + Field + """: """ + value + """}]}"
    
    ' 4. Envoyer la mise à jour
    url = "https://api.jsonbin.io/v3/b/" & BinId
    With http
        .Open "PUT", url, False
        .setRequestHeader "X-Master-Key", JSONBIN_API_KEY
        .setRequestHeader "Content-Type", "application/json"
        .send updatedJson
        If .Status = 200 Or .Status = 201 Then
            SetValueToBin = True
        Else
            SetValueToBin = False
        End If
    End With
End Function

