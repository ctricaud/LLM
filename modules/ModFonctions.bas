Attribute VB_Name = "ModFonctions"
Option Explicit
Sub NettoyerBorduresHorizontalesVisibles()

    Dim wn As Window
    Dim ws As Worksheet
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim r As Long
    Dim rngRow As Range
    Dim hasTop As Boolean, hasBottom As Boolean
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Sheets
        
       

            firstCol = 1
            lastCol = 1000
            
                For r = 1 To 10
                    
                    Set rngRow = ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol))
                    
             
                        rngRow.Borders(xlEdgeTop).LineStyle = xlNone
                 
                        rngRow.Borders(xlEdgeBottom).LineStyle = xlNone

                    
                Next r
                
   
            

        
Next ws
    
    Application.ScreenUpdating = True
    
End Sub


Private Function BordureContinue(rng As Range, edgeType As XlBordersIndex) As Boolean
    
    Dim c As Range
    
    For Each c In rng.Cells
        If c.Borders(edgeType).LineStyle = xlNone Then
            BordureContinue = False
            Exit Function
        End If
    Next c
    
    BordureContinue = True
    
End Function

Public Sub ExportModules()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String, fName As String
    Dim wbPath As String, modulesPath As String
    Dim oShell As Object
    
    ' Répertoire du classeur
    wbPath = ThisWorkbook.Path
    If wbPath = "" Then
        MsgBox "Le fichier doit être enregistré avant d'exécuter cette macro.", vbExclamation
        Exit Sub
    End If
    
    ' Répertoire Modules
    modulesPath = CheminFichier & "\Modules"
    If dir(modulesPath, vbDirectory) = "" Then MkDir modulesPath
    
    ' Export des modules
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_Document
                fName = modulesPath & "\" & vbComp.Name & ExtensionModule(vbComp.Type)
                On Error Resume Next
                Kill fName ' Supprime si existe
                On Error GoTo 0
                vbComp.Export fName
        End Select
    Next vbComp
    
    ' Création du ZIP
    Exit Sub
    ' Création du fichier ZIP
    Dim ZipPath
    ZipPath = CheminFichier & "\Modules.zip"
    If dir(ZipPath) <> "" Then Kill ZipPath
    
    ' Crée un ZIP vide
    Open ZipPath For Output As #1
    Print #1, "PK" & Chr$(5) & Chr$(6) & String(18, vbNullChar)
    Close #1
    
    ' Instanciation correcte de l'objet Shell
    Set oShell = CreateObject("Shell.Application")
    Dim oZip As Object
    Set oZip = oShell.Namespace(ZipPath)
    
    ' Ajout des fichiers exportés dans le zip
    oZip.CopyHere oShell.Namespace(modulesPath).Items
    
    ' Attente (sinon copie incomplète si fichiers nombreux ou gros)
    Application.Wait (Now + TimeValue("0:00:02"))
    
    MsgBox "Modules exportés et compressés dans : " & ZipPath
End Sub

Private Sub CreerZipVide(cheminZip As String)
    Dim numFichier As Integer
    
    numFichier = FreeFile
    Open cheminZip For Output As numFichier
    Print #numFichier, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close numFichier
End Sub
Private Function ExtensionModule(vbType As VBIDE.vbext_ComponentType) As String
    Select Case vbType
        Case vbext_ct_ClassModule: ExtensionModule = ".cls"
        Case vbext_ct_StdModule: ExtensionModule = ".bas"
        Case vbext_ct_MSForm: ExtensionModule = ".frm"
        Case vbext_ct_Document: ExtensionModule = ".cls"
        Case Else: ExtensionModule = ".txt"
    End Select
End Function
Sub LogCA12Mois()
    '-------------------------------------------------------------------------------------
    'Cette fonction permet d'ajouter aux log les informations sur les CA 12 Mois
    '-------------------------------------------------------------------------------------
    Dim CA12Mois As Variant
    CA12Mois = Range("HistoriqueCA").ListObject.DataBodyRange.value
    
    log ""
    log "-------------------------------------"
    log "D. Historique CA et Stock"
    log "-------------------------------------"
    log "CA 12 Mois Lofts : " + FormatNumber(CA12Mois(1, 2), 2) + "€"
    log "CA 12 Mois Total : " + FormatNumber(CA12Mois(1, 3), 2) + "€"
    log "Stock Lofts : " + FormatNumber(CA12Mois(1, 4), 2) + "€"
    log "Stock Lofts N-1 : " + FormatNumber(CA12Mois(1, 5), 2) + "€"
    
    
End Sub


Function TrouverFeuilleDuTableau(Tableau As String) As String
    Dim ws As Worksheet
    Dim lo As ListObject
    
    ' Parcourir toutes les feuilles du classeur
    For Each ws In ThisWorkbook.Worksheets
        ' Parcourir les tableaux structurés de la feuille
        For Each lo In ws.ListObjects
            ' Vérifier si le nom du tableau correspond
            If lo.Name = Tableau Then
                ' Retourner la feuille contenant le tableau
                TrouverFeuilleDuTableau = ws.Name
                Exit Function
            End If
        Next lo
    Next ws
    
    ' Si le tableau n'est pas trouvé, renvoyer Nothing
    TrouverFeuilleDuTableau = ""
End Function

Function ConvertirDate(DateInput As Variant, Optional SeparateurRetour As String = "") As String
    Dim Sep As String
    Dim Part1 As String, Part2 As String, Part3 As String
    
    ' Détecte le séparateur utilisé dans la date d'entrée
    If InStr(DateInput, "-") > 0 Then
        Sep = "-"
    ElseIf InStr(DateInput, "/") > 0 Then
        Sep = "/"
    ElseIf InStr(DateInput, ".") > 0 Then
        Sep = "."
    Else
        ConvertirDate = "Format incorrect"
        Exit Function
    End If
    
    ' Si SeparateurRetour n'est pas fourni, utiliser le séparateur d'origine
    If SeparateurRetour = "" Then
        SeparateurRetour = Sep
    End If
    
    ' Sépare les parties de la date
    Part1 = Split(DateInput, Sep)(0)
    Part2 = Split(DateInput, Sep)(1)
    Part3 = Split(DateInput, Sep)(2)
    
    ' Vérifie le format et transforme en conséquence
    If Len(Part1) = 4 Then
        ' Format yyyy-mm-dd à dd-mm-yyyy
        ConvertirDate = Part3 & SeparateurRetour & Part2 & SeparateurRetour & Part1
    ElseIf Len(Part3) = 4 Then
        ' Format dd-mm-yyyy à yyyy-mm-dd
        ConvertirDate = Part3 & SeparateurRetour & Part2 & SeparateurRetour & Part1
    Else
        ConvertirDate = "Format incorrect"
    End If
End Function

Function GetClipboardContent() As String
    Dim MyData As DataObject
    Set MyData = New DataObject
    On Error Resume Next
    MyData.GetFromClipboard
    GetClipboardContent = MyData.GetText(1)
    On Error GoTo 0
End Function
Function ConvertirEncodage(ByVal texte) As String
    ' Remplace les séquences de caractères mal encodés par les bons caractères
    ' Ce sont les caractères que vous avez mentionnés dans votre exemple
    texte = Replace(texte, "ÃƒÂ©", "é")
    texte = Replace(texte, "Ã©", "é")
    texte = Replace(texte, "ÃƒÂ¨", "è")
    texte = Replace(texte, "ÃƒÂª", "ê")
    texte = Replace(texte, "Â ", "/")
    texte = Replace(texte, "ÃƒÂ´", "ô")
    texte = Replace(texte, "Ã‚Â ", " ")
    texte = Replace(texte, "Ãƒ", "f") ' selon le contexte cela pourrait changer
    texte = Replace(texte, "Ã‚", "")
    
    ' Ajoutez d'autres remplacements si nécessaire pour traiter tous les caractères mal encodés

    ConvertirEncodage = texte
End Function
Function DateUS(D As Date) As String
'----------------------------------------------------
'Cette procédure calcule la date en anglais
'----------------------------------------------------
Dim months As Variant
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
   DateUS = CStr(Day(D)) + " " + CStr(months(Month(Date) - 1)) + " " + CStr(Year(D))

End Function


Function ImporterCSVDansTableau(CheminFichier)
    Dim csvFeuille As Worksheet
    Dim derniereligne As Long, dernierecolonne As Long
    Dim i As Long, j As Long
    
    ' Ouvrir le fichier CSV en tant que feuille Excel avec Workbooks.OpenText
    Workbooks.OpenText Filename:=CheminFichier, DataType:=xlDelimited, Comma:=True
    
    ' Référence à la feuille active où le CSV est importé
    Set csvFeuille = ActiveSheet
    
    ' Trouver la dernière ligne et dernière colonne de données
    derniereligne = csvFeuille.Cells(csvFeuille.Rows.Count, 1).End(xlUp).Row
    dernierecolonne = csvFeuille.Cells(1, csvFeuille.Columns.Count).End(xlToLeft).Column
    
    ' Copier les données dans un tableau
    ImporterCSVDansTableau = csvFeuille.Range(csvFeuille.Cells(1, 1), csvFeuille.Cells(derniereligne, dernierecolonne)).value
    
    ' Fermer le fichier CSV sans sauvegarder les modifications
    csvFeuille.Parent.Close SaveChanges:=False
    
    ' Libérer les variables
    Set csvFeuille = Nothing
End Function

Function ExtraitNombre(texte As String) As String
    Dim iPos As Long
    Dim Extrait As String
    
    For iPos = 1 To Len(texte)
        If InStr("0123456789.,", Mid(texte, iPos, 1)) <> 0 Then
            Extrait = Extrait + Mid(texte, iPos, 1)
        End If
    Next iPos
    
    ExtraitNombre = Extrait
    
End Function


Function RECHERCHEV(Valeur_Cherchee As Variant, Table_matrice As Range, No_index_col As Single, Optional Valeur_proche As Boolean)
'par Excel-Malin.com ( https://excel-malin.com/ )
 
On Error GoTo RECHERCHEVerror
    RECHERCHEV = Application.VLookup(Valeur_Cherchee, Table_matrice, No_index_col, Valeur_proche)
    If IsError(RECHERCHEV) Then RECHERCHEV = "#N/A"
    
Exit Function
RECHERCHEVerror:
    RECHERCHEV = "#N/A"
End Function
Function LireFichierTexte(nomFichier As String) As Variant
    Dim objFso As Object
    Dim objFichier As Object
    Dim lignes() As String
    Dim ligneCourante As String
    Dim i As Integer

    ' Création de l'objet FileSystemObject
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' Vérification si le fichier existe
    If objFso.FileExists(nomFichier) Then
        ' Récupérer l'objet fichier
        Set objFichier = objFso.GetFile(nomFichier)
        
        ' Écrire la date de dernière modification dans la fenêtre de débogage
        log "Date du fichier Prix AirBNB: " & objFichier.DateLastModified
        
        ' Ouvrir le fichier pour lecture
        Dim objTexte As Object
        Set objTexte = objFso.OpenTextFile(nomFichier, 1)
        
        ' Lire les lignes et les stocker dans un tableau
        i = 0
        Do While Not objTexte.AtEndOfStream
            ligneCourante = objTexte.ReadLine
            ReDim Preserve lignes(i)
            lignes(i) = ligneCourante
            i = i + 1
        Loop
        
        ' Fermer le fichier
        objTexte.Close
        
        ' Renvoyer le tableau des lignes
        LireFichierTexte = lignes
    Else
        ' Si le fichier n'existe pas, renvoyer un message d'erreur
        'Debug.Print "Erreur : Le fichier n'existe pas."
        LireFichierTexte = Array()
    End If
    
    ' Libération des objets
    Set objTexte = Nothing
    Set objFichier = Nothing
    Set objFso = Nothing
End Function

Function MoisUSDate(D) As String
    'remplace 12 feb 2024 par 12/02/2024
    D = Replace(D, " Jan ", "/01/")
    D = Replace(D, "Feb ", "/02/")
    D = Replace(D, " Mar ", "/03/")
    D = Replace(D, " Apr ", "/04/")
    D = Replace(D, " May ", "/05/")
    D = Replace(D, " Jun ", "/06/")
    D = Replace(D, " Jul ", "/07/")
    D = Replace(D, " Aug ", "/08/")
    D = Replace(D, " Sep ", "/09/")
    D = Replace(D, " Oct ", "/10/")
    D = Replace(D, " Nov ", "/11/")
    D = Replace(D, " Dec ", "/12/")
    
    MoisUSDate = D
End Function

Sub FormatDatesAsShortDate()
  
    Dim ws As Worksheet
    Dim cell As Range
    
    ' Boucle à travers toutes les feuilles du classeur
    
    'For Each ws In ThisWorkbook.Worksheets
        Set ws = ActiveSheet
        ws.Unprotect
        ' Boucle à travers chaque cellule dans la feuille
        For Each cell In ws.UsedRange
            ' Vérifie si la cellule contient une date
            If IsDate(cell.value) Then
                ' Change le format de la cellule en date courte
                cell.NumberFormat = "dd/mm/yyyy"
            ElseIf IsNumeric(cell.value) Then
                If InStr(1, cell.NumberFormat, "$") > 0 Then
                    cell.NumberFormat = "$* # ##0.00;$* -# ##0.00"
                End If
            End If
        Next cell
    'Next ws
    ws.Unprotect
    'MsgBox Timer * 1000 - dd
End Sub

Function ISO8601ToDate(isoDate) As Date
    Dim dt As String
    Dim D As Date
    
    ' Retire le Z s'il existe
    If Right(isoDate, 1) = "Z" Then
        isoDate = Left(isoDate, Len(isoDate) - 1)
    End If
    
    ' Remplace le T par un espace
    dt = Replace(isoDate, "T", " ")
    
    ' Retire la partie millisecondes si elle existe
    If InStr(dt, ".") > 0 Then
        dt = Left(dt, InStr(dt, ".") - 1)
    End If
    
    ' Conversion en date
    On Error Resume Next
    D = CDate(dt)
    On Error GoTo 0
    
    ISO8601ToDate = D
End Function



Function CheminFichier() As String
    'Cette précodéure permet de récupérer le nom du fichier excel actif
    Dim baseracineonedrive
    Dim Chemin
    Dim c As Integer
    Dim old_Chemin As String
    
    baseracineonedrive = Environ("OneDrive") & "\"
    Chemin = ActiveWorkbook.Path
    c = InStr(1, Chemin, ".net/")
    c = InStr(c + 6, Chemin, "/")
    old_Chemin = Left(Chemin, c)
    CheminFichier = Replace(Replace(Chemin, old_Chemin, baseracineonedrive), "/", "\") & "\"
End Function

Sub RAZFiltres(Tableau As String)
    'Dim ws As String
    'ws = TrouverFeuilleDuTableau(Tableau)
    'If ws <> "" Then
    With Range(Tableau).ListObject
        If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
        .ShowAutoFilter = True
    End With
    'End If
End Sub




Sub RAZTableau(Tableau As String, Optional nbLignes = 0)
    'Dim ws As String
    'ws = TrouverFeuilleDuTableau(Tableau)
    'If ws <> "" Then
    If Not Range(Tableau).ListObject.DataBodyRange Is Nothing Then
        Range(Tableau).ListObject.DataBodyRange.Delete
    End If
    
    If nbLignes > 0 Then
        Dim tListObject As ListObject
        Set tListObject = Range(Tableau).ListObject
        tListObject.Resize tListObject.Range.Resize(nbLignes + 1)
    End If
    'End If
End Sub


Sub Sauvegarde()
    '============================================
    'Effectue une sauvegarde du fichier excel dans le répertoire Sauvegarde
    '============================================
    Dim nomFichier As String

    nomFichier = CheminFichier() + "Sauvegarde"
    'On vérifie si le répertoire Export existe, sinon on le crée
    If Len(dir(nomFichier, vbDirectory)) = 0 Then
        MkDir (nomFichier)
    End If

    nomFichier = nomFichier + "\Sauvegarde Suivi réservations" + " - " + Replace(Replace(CStr(Now()), "/", "-"), ":", "-") + ".xlsm"

    ThisWorkbook.SaveCopyAs nomFichier
End Sub
Function nbJoursMois(mois, annee)
    'Retourne le nombre de jours d'un mois passé en paramètre
    Dim dd As String
    
    dd = "01/" + CStr(mois) + "/" + CStr(annee)
    nbJoursMois = DateAdd("m", 1, dd) - CDate(dd)
End Function
Function LettresColonne(NoCol)
    'Retourne le numéro de colonne à partir des lettres Excel
    LettresColonne = Split(Cells(1, NoCol).Address, "$")(1)
End Function
Function indexRange(Nom, colonne, valeur) As Integer
   'Retourne l'index d'une valeur dans un tableau
   Dim i As Integer
   Dim r As Range
   
    Set r = Range(Nom)
    indexRange = 0      'Valeur si pas trouvé
    
    For i = 1 To r.ListObject.ListRows.Count
        If r(i, colonne).value = valeur Then
            indexRange = i
            Exit For
        End If
    Next i
End Function
Public Function BuildIndexCache(ByVal nomPlage As String, ByVal col As Long) As Object
    Dim lo As ListObject
    Dim arr As Variant
    Dim D As Object
    Dim i As Long
    
    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = vbTextCompare  ' recherche insensible à la casse
    
    On Error GoTo EH
    Set lo = Range(nomPlage).ListObject
    arr = lo.DataBodyRange.value
    
    For i = 1 To UBound(arr, 1)
        D(CStr(arr(i, col))) = i
    Next i
    
    Set BuildIndexCache = D
    Exit Function
EH:
    ' Si le tableau est vide, renvoie dict vide
    Set BuildIndexCache = D
End Function
Sub SupprimerFichiers(Nom As String)
    Dim cheminDossier As String
    Dim fichier As String
    Dim fso As Object
    
    ' Récupérer le chemin du répertoire depuis la cellule "Download"
    cheminDossier = Range("DirDownload").value
    If Right(cheminDossier, 1) <> "\" Then cheminDossier = cheminDossier & "\"
    
    ' Créer l'objet FileSystemObject pour manipuler les fichiers
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Parcourir tous les fichiers correspondant à *statement*.csv dans le répertoire
    fichier = dir(cheminDossier & Nom)
    
    Do While fichier <> ""
        ' Supprimer chaque fichier trouvé
        fso.DeleteFile cheminDossier & fichier
        ' Passer au fichier suivant
        fichier = dir
    Loop
    
    ' Libérer l'objet FileSystemObject
    Set fso = Nothing
    
End Sub


Sub TableToTableau(T As Variant, TableName As String, Optional nbLignes As Long = -1)
    '------------------------------------------------------
    'Cette proécdure permet de transférer une array dans un tableau
    '------------------------------------------------------
    '1. On efface les filtres
    '------------------------
    RAZFiltres TableName
    
    '2. On remet à zéro le tableau
    '------------------------
    If nbLignes = -1 Then nbLignes = (UBound(T) - LBound(T) + 1)
    RAZTableau TableName, nbLignes
    
        
    '4. On recopie la table dans le tableau vierge
    '-------------------------------------
    Range(TableName).ListObject.DataBodyRange.value = T
    
End Sub

