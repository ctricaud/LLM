Attribute VB_Name = "ModSQL"

Option Explicit

Sub Exemple()
    Dim strSql As String
    Dim TableSQL As Variant
    
    '-------------------------------------------------------------------
    '1. On écrit et exécute exécute la requête
    '-------------------------------------------------------------------
    'strSql = "SELECT table1.Logement, " _
            + "table2.idFoncier, " _
            + "SUM(IIF(IsNumeric([Montant]), [Montant], 0)) as TotalMontant, " _
            + "SUM(IIF(IsNumeric([Montant]), IIF([Montant]<0,-[Montant], 0),0)) as TotalHost, " _
            + "SUM(IIF(IsNumeric([Frais de ménage]), [Frais de ménage], 0)) as TotalMenage, " _
            + "SUM(IIF(IsNumeric([Frais de service]), [Frais de Service], 0)) as TotalService, " _
            + "TotalMontant+TotalService+TotalHost as Total " _
            + "FROM " + Plage("Tableau") + " as table1 " _
            + "INNER JOIN " + Plage("Logements") + " as table2 " _
            + "ON table1.logement = table2.logement " _
            + "WHERE [Année des revenus] = " + CStr(Range("Annee")) + " " _
           + "GROUP BY table1.Logement,idFoncier"
          
    strSql = "SELECT Ménage,Comm FROM " + plage("Logements") + " as table1 WHERE  table1.Logements ='Maury'"
    TableSQL = ExecuteSQL(strSql)(0, 0)
    
    '-------------------------------------------------------------------
    '2. On met à jour le tableau
    '-------------------------------------------------------------------
    'TableToTableau TableSQL, "Synthese"

End Sub



Function ExecuteSQL(Requete As String) As Variant
    Dim TableauExcel As DAO.Database
    Dim rs As DAO.Recordset
    Dim ListePrenoms As String
    Dim strSql


    '============================================================
    'Etape 1 : création du code SQL
    '============================================================
    strSql = Requete

    '============================================================
    'Etape 2 : Execution du code et récupération dans un RecordSet
    '============================================================
    Set TableauExcel = OpenDatabase(CheminFichier + ThisWorkbook.Name, False, False, "Excel 8.0")

    'Exécuter la requête
    Set rs = TableauExcel.OpenRecordset(strSql)

    '============================================================
    'Etape 3 : Création de la liste déroulante
    '============================================================
    rs.MoveFirst
    ExecuteSQL = TransposeArray(rs.GetRows(rs.RecordCount))
    
    'Fermeture de la base de données
    rs.Close
    Set rs = Nothing
    TableauExcel.Close
    Set TableauExcel = Nothing

End Function

Function GetSheetNameOfTable(TableName) As String

    Dim tbl As ListObject
    Dim sheetName As String
    Dim ws As Worksheet
        
    'Parcourir toutes les feuilles pour trouver le tableau
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(TableName)
        On Error GoTo 0
        
        ' Vérifier si le tableau a été trouvé sur la feuille
        If Not tbl Is Nothing Then
            sheetName = ws.Name
            Exit For
        End If
    Next ws
    
    ' Afficher le nom de la feuille
    If sheetName <> "" Then
        GetSheetNameOfTable = sheetName
    Else
        GetSheetNameOfTable = ""
    End If
End Function





Function plage(TableName) As String
    '1. On récupère la feuille
    Dim sheetName As String
    sheetName = GetSheetNameOfTable(TableName)
    
    '2. On récupère l'adresse complète
    Dim adresse As String
    adresse = Replace(ThisWorkbook.Sheets(sheetName).ListObjects(TableName).Range.Address, "$", "")
    
    '3. On compose la plage
    plage = "[" + sheetName + "$" + adresse + "]"
    
End Function


Function TransposeArray(T As Variant) As Variant
    Dim i As Long, j As Long
    Dim minRow As Long, maxRow As Long
    Dim minCol As Long, maxCol As Long
    Dim Transposed As Variant
    
    ' Obtenir les bornes des lignes et des colonnes
    minRow = LBound(T, 1)
    maxRow = UBound(T, 1)
    minCol = LBound(T, 2)
    maxCol = UBound(T, 2)
    
    ' Redimensionner le tableau transposé avec les bornes inversées
    ReDim Transposed(minCol To maxCol, minRow To maxRow)
    
    ' Transposer le tableau
    For i = minRow To maxRow
        For j = minCol To maxCol
            Transposed(j, i) = T(i, j)
        Next j
    Next i
    
    ' Renvoyer le tableau transposé
    TransposeArray = Transposed
End Function

