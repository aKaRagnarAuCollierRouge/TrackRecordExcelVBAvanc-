Attribute VB_Name = "FindIndexCollumn"


Function FindIndiceCollumnWithWorkSheetName(sheetName As String, CollName As String) As Integer
    Dim ws As Worksheet
    Dim lastColumn As Integer
    Dim col As Integer
    Dim foundColumn As Boolean
    
    Set ws = Worksheets(sheetName)
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    foundColumn = False
    
    ' Boucle à travers les colonnes pour trouver les en-têtes
    For col = 1 To lastColumn
        If ws.Cells(1, col).value = CollName Then
            ' Retourner l'indice de la colonne trouvée
            FindIndiceCollumn = col
            foundColumn = True
            Exit For ' Sortir de la boucle dès que la colonne est trouvée
        End If
    Next col
    
    If Not foundColumn Then
        ' Retourner 0 si la colonne n'est pas trouvée
        FindIndiceCollumn = 0
    End If
End Function

Function FindIndiceCollumnWithTable(ws As Worksheet, tableName As String, colName As String) As Integer
    Dim tbl As ListObject
    Dim headerRow As ListRow
    Dim foundColumn As Boolean
    Dim col As Integer
    
    ' Initialiser la fonction à 0 (si la colonne n'est pas trouvée)
    FindIndiceCollumnWithTable = 0
    foundColumn = False
    
    ' Obtenir l'objet tableau
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "Le tableau '" & tableName & "' n'existe pas dans la feuille " & ws.Name
        Exit Function
    End If
    
    ' Parcourir les noms de colonnes pour trouver l'indice
    For col = 1 To tbl.ListColumns.Count
        If StrComp(tbl.ListColumns(col).Name, colName, vbTextCompare) = 0 Then
            FindIndiceCollumnWithTable = col
            foundColumn = True
            Exit For
        End If
    Next col
    
    If Not foundColumn Then
        Debug.Print "Colonne '" & colName & "' non trouvée dans le tableau '" & tableName & "' de la feuille " & ws.Name
    End If
End Function

Sub TestFindIndiceCollumnWithTable()
    Dim ws As Worksheet
    Dim colIndex As Integer
    Dim tableName As String
    Set ws = ThisWorkbook.Worksheets("Trackrecord")
    tableName = "Tableau1" ' Remplacez par le nom réel de votre tableau
    
    colIndex = FindIndiceCollumnWithTable(ws, tableName, "Date Début")
    If colIndex > 0 Then
        Debug.Print "Colonne 'Date Début' trouvée à l'indice : " & colIndex
    Else
        Debug.Print "Colonne 'Date Début' non trouvée."
    End If
End Sub


Function FindColumnsWithPattern(ws As Worksheet, tableName As String, pattern As String) As Collection
    Dim tbl As ListObject
    Dim colIndices As Collection
    Dim col As Integer
    
    ' Initialiser la collection pour stocker les indices de colonnes correspondantes
    Set colIndices = New Collection
    
    ' Obtenir l'objet tableau
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Debug.Print "Le tableau '" & tableName & "' n'existe pas dans la feuille " & ws.Name
        Exit Function
    End If
    
    ' Parcourir les noms de colonnes pour trouver les colonnes contenant le motif
    For col = 1 To tbl.ListColumns.Count
        If InStr(1, tbl.ListColumns(col).Name, pattern, vbTextCompare) > 0 Then
            colIndices.Add col
        End If
    Next col
    
    ' Retourner la collection des indices de colonnes trouvées
    Set FindColumnsWithPattern = colIndices
End Function

Sub TestFindColumnsWithPattern()
    Dim ws As Worksheet
    Dim colIndices As Collection
    Dim tableName As String
    Dim col As Variant
    
    ' Définir la feuille et le nom du tableau
    Set ws = ThisWorkbook.Worksheets("TrackRecord")
    tableName = "Tableau1" ' Remplacez par le nom réel de votre tableau
    
    ' Appeler la fonction pour trouver les colonnes avec le motif "Screenshot"
    Set colIndices = FindColumnsWithPattern(ws, tableName, "Screenshot")
    
    ' Afficher les indices des colonnes trouvées
    If colIndices.Count > 0 Then
        For Each col In colIndices
            Debug.Print "Colonne trouvée à l'indice : " & col
        Next col
    Else
        Debug.Print "Aucune colonne contenant le motif 'Screenshot' trouvée."
    End If
End Sub
