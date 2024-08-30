Attribute VB_Name = "TrieEtPreparation"
Function FiltreRowByDate(DateDebut As Date, DateFin As Date, ColumnDate As Integer, WorksheetName As String) As Collection
    Dim TrackRecord As Worksheet
    Dim Rapport As Worksheet
    Dim lignesFiltrees As Collection
    Dim ligneArray() As Variant
    Dim lastRow As Integer
    Dim row As Long  ' Déclarer i
    Set TrackRecord = Worksheets(WorksheetName)
    lastRow = TrackRecord.Cells(TrackRecord.Rows.Count, 1).End(xlUp).row
    
    Set lignesFiltrees = New Collection
    
    For row = 1 To lastRow
        ' Vérifier si la date dans la colonne A est comprise entre dateDebut et dateFin
        If IsDate(TrackRecord.Cells(row, ColumnDate).value) Then
            If TrackRecord.Cells(row, ColumnDate).value >= DateDebut And TrackRecord.Cells(row, ColumnDate).value <= DateFin Then
                ' Stocker la ligne entière dans la collection
                ligneArray = TrackRecord.Rows(row).value
                lignesFiltrees.Add ligneArray
            End If
        End If
    Next row
    ' Retourner la collection
    Set FiltreRowByDate = lignesFiltrees
End Function
