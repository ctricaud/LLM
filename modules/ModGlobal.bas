Attribute VB_Name = "ModGlobal"
'+++++++++++++++++++++++++++++++++++++++++++++++
'+ Ce module présente des fonctions transverses
'+++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Sub WaitReadyState(driver As WebDriver)
Dim ReadyState
    Do
    ReadyState = driver.ExecuteScript("return document.readyState")
    ' Attendre une courte période avant de vérifier à nouveau
    Application.Wait Now + TimeValue("00:00:01")
Loop While ReadyState <> "complete"
End Sub
Sub ApplicationOntime(DateProc As Date, Procedure As String)
    '----------------------------------------------------------
    'Lancement différé de procédure avec suivi
    '----------------------------------------------------------
    '1. On rajoute la planification
    '----------------------------------------------------------
    Dim NewRow As Object

    Set NewRow = Feuil11.Range("Planificateur").ListObject.ListRows.Add
    NewRow.Range(1, 1) = Now
    NewRow.Range(1, 2) = DateProc
    NewRow.Range(1, 3) = Procedure

    '----------------------------------------------------------
    '1. On trie le tableau
    '----------------------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Accueil")
    Set tbl = ws.ListObjects("Planificateur")

    ' Trier le tableau structuré par la seconde colonne (Date) en ordre croissant
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(2).DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .header = xlYes
        .Apply
    End With
    
    '----------------------------------------------------------
    '3. On lance la proécdure
    '----------------------------------------------------------
    Application.OnTime DateProc, Procedure
    

End Sub
Sub ControlePlanificateur(Procedure)
    '----------------------------------------------------------
    'Cette procédure supprime la présence dans le planificateur
    '----------------------------------------------------------
    '1 On cherche dans le planificateur la procédure
    '----------------------------------------------------------
    Dim i As Integer
    
    For i = 1 To UBound(Feuil11.Range("Planificateur").value)
        If Feuil11.Range("Planificateur[Procédure]")(i) = Procedure Then
            'If Abs(DateDiff("s", Now, Range("Planificateur[Exécution]")(i))) < 3 Then
                Range("Planificateur").ListObject.ListRows(i).Delete
            'End If
        End If
    Next i
    
    
    
End Sub


Function GetTablePrix()
    '---------------------------------------------------
    'Cette fonction permet de récupérer la liste des prix
    'Pour tous les logements, mis sur une feuille de calcul
    '---------------------------------------------------
    '1. On récupère les information de l'extraction web
    '---------------------------------------------------
    Dim T As Variant
    T = Range("TableauPrix").value
    Dim dateDebut As Date: dateDebut = T(LBound(T), 1)
    Dim dateFin As Date: dateFin = T(UBound(T), 1)
    'dateFin = DateAdd("d", 364, Date)
    Dim nbDates As Integer: nbDates = UBound(T) - LBound(T) + 1
    'nbDates = 364
    Dim nbLogements: nbLogements = UBound(T, 2) - LBound(T, 2)
     
    '--------------------------------------------------------------
    '2. On retranscrit les prix téléchargés pour pouvoir mieux les exploiter
    '----------------------------------------------------------------
    Dim TablePrix As Variant
    
    ReDim TablePrix(nbLogements, dateDebut - 1 To dateDebut + nbDates - 1)
    Dim iLogement As Integer
    Dim iDate As Long
    
    For iLogement = 1 To nbLogements
        For iDate = 1 To nbDates
            TablePrix(iLogement - 1, iDate + dateDebut - 1) = T(iDate, iLogement + 1)
        Next iDate
    Next iLogement
 
    GetTablePrix = TablePrix
End Function







Sub log(texte)
    If Left(texte, 1) = "=" Then texte = "'" + texte
    Feuil8.Range("LogExtraction") = Range("LogExtraction") + texte + vbCrLf
    DoEvents
End Sub


Sub Notification(DateNotification As Date, channel As String, titreNotification As String, typeNotification As String, Source As String)
    '-------------------------------------------------------------
    'Cette procédure permet d'ajouter une notification dans
    'le TS Notifs
    '-------------------------------------------------------------
    Dim U As Object
    
    Set U = Range("Notifs").ListObject.ListRows.Add(1)
          
    U.Range(1, 1) = Format(Now, "dd-mm-yy hh:nn")
    U.Range(1, 2) = channel
    U.Range(1, 3) = typeNotification
    U.Range(1, 4) = "'" + titreNotification

    U.Range(1, 5) = Source
    U.Range(1, 6) = "X"
End Sub

Sub Todo(texte)
   'On remplit le carré todo pour qu'il reste permanent
   log (texte)
    If Range("Todo") <> "" Then Range("Todo") = Range("Todo") + vbCrLf + vbCrLf
    Range("Todo") = Range("Todo") + Format(Now, "dd-mm-yyyy hh:mm:ss") + vbCrLf + texte + vbCrLf
    DoEvents
End Sub


Sub Initialisation()
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+ Cette préocédure est lancée à l'ouverture du workbook
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'Les variables
    
    
    'Les procédures
    CA12Mois
    
End Sub


