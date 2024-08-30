Attribute VB_Name = "PreparationDatasDayResume"
Public Function PreparationResumeDay(DateJour As Date) As Collection
    Dim RowFiltred As Collection
    Dim wsTrackrecord As Worksheet
    Dim tableName As String
    Dim ColonneRR As Integer
    Dim ColonneDateDebut As Integer
    Dim ColonneDateFin As Integer
    Dim ColonneHeureEntree As Integer
    Dim ColonneHeureSortie As Integer
    Dim ColonneKeyTrade As Integer
    Dim ColonneGain As Integer
    Dim ColonneActif As Integer
    Dim JSONdico As Object
    Dim dicoDay As Scripting.Dictionary
    Dim resultCollection As Collection
    Dim screenshotCols As Collection
    Dim i As Integer
    Dim winCount As Integer
    Dim lossCount As Integer
    Dim totalRR As Double
    
    tableName = "Tableau1" ' Remplacez par le nom réel de votre tableau
    Set wsTrackrecord = ThisWorkbook.Worksheets("TrackRecord")
    Set RowFiltred = FiltreRowByDate(DateJour, DateJour, 1, "TrackRecord")
    ColonneDateFin = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date Fin")
    ColonneRR = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "RR")
    ColonneGain = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Gain")
    ColonneDateDebut = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date Début")
    ColonneHeureEntree = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure Début")
    ColonneHeureSortie = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure Fin")
    ColonneKeyTrade = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "KeyTrade")
    ColonneActif = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Actif")
    
    ' Obtenir la collection des colonnes "Screenshot"
    Set screenshotCols = FindColumnsWithPattern(wsTrackrecord, tableName, "Screenshot")
    
    ' Charger le fichier JSON et obtenir le dictionnaire pour le jour spécifique
    Set JSONdico = LoadjsonObject(year(DateJour))
    Set dicoDay = GetDictionnaryDayCommentary(JSONdico, month(DateJour), day(DateJour))
    
    ' Initialiser les collections et les variables de comptage
    Set resultCollection = New Collection
    winCount = 0
    lossCount = 0
    totalRR = 0
    
    ' Boucler sur les lignes filtrées
    For i = 1 To RowFiltred.Count
        Dim rowData As Scripting.Dictionary
        Dim ligneArray As Variant
        Dim rrValue As Double
        Dim keyTrade As String
        Dim Commentary As String
        
        ligneArray = RowFiltred(i)
        
        ' Créer un nouveau dictionnaire pour stocker les données de la ligne
        Set rowData = New Scripting.Dictionary
        rowData.Add "Date Début", ligneArray(1, ColonneDateDebut)
        rowData.Add "Date Fin", ligneArray(1, ColonneDateFin)
        rowData.Add "Heure Entree", ligneArray(1, ColonneHeureEntree)
        rowData.Add "Heure Sortie", ligneArray(1, ColonneHeureSortie)
        rowData.Add "RR", ligneArray(1, ColonneRR)
        rowData.Add "Gain", ligneArray(1, ColonneGain)
        rowData.Add "KeyTrade", ligneArray(1, ColonneKeyTrade)
        rowData.Add "Actif", ligneArray(1, ColonneActif)
        
        ' Ajouter les valeurs des colonnes "Screenshot" au dictionnaire
        Dim colIndex As Variant
        For Each colIndex In screenshotCols
            rowData.Add "Screenshot" & colIndex, ligneArray(1, colIndex)
        Next colIndex
        
        rrValue = ligneArray(1, ColonneRR)
        keyTrade = ligneArray(1, ColonneKeyTrade)
        
        ' Obtenir le commentaire associé au Key Trade
        If dicoDay.Exists(keyTrade) Then
            Commentary = dicoDay(keyTrade)("Commentary")
        Else
            Commentary = ""
        End If
        rowData.Add "Commentary", Commentary
        
        ' Compter les wins et les losses
        If rrValue > 0 Then
            winCount = winCount + 1
        ElseIf rrValue < 0 Then
            lossCount = lossCount + 1
        End If
        
        ' Ajouter le RR au total
        totalRR = totalRR + rrValue
        
        ' Ajouter le dictionnaire à la collection de résultats
        resultCollection.Add rowData
    Next i
    
    ' Afficher les résultats pour vérification
    Debug.Print "Wins: " & winCount
    Debug.Print "Losses: " & lossCount
    Debug.Print "Total RR: " & totalRR
    
    For Each rowData In resultCollection
        Debug.Print "Date Début: " & rowData("Date Début")
        Debug.Print "Date Fin: " & rowData("Date Fin")
        Debug.Print "Heure Entree: " & rowData("Heure Entree")
        Debug.Print "Heure Sortie: " & rowData("Heure Sortie")
        Debug.Print "RR: " & rowData("RR")
        Debug.Print "Gain: " & rowData("Gain")
        Debug.Print "Key Trade: " & rowData("KeyTrade")
        Debug.Print "Commentary: " & rowData("Commentary")
        Debug.Print "Actif: " & rowData("Actif")
        For Each colIndex In screenshotCols
            Debug.Print "Screenshot" & colIndex & ": " & rowData("Screenshot" & colIndex)
        Next colIndex
        Debug.Print "-------------------------------"
    Next rowData
    
    Set PreparationDatasResumeDay = resultCollection
End Function

Sub TestPreparationDatasResumeDay()
    Dim result As Collection
    Dim rowData As Scripting.Dictionary
    Dim testDate As Date
    
    ' Définir une date pour le test
    testDate = #1/1/2024#
    
    ' Appeler la fonction principale avec la date de test
    Set result = PreparationDatasResumeDay(testDate)
    
    ' Vérifier et afficher les résultats dans la fenêtre de débogage
    Debug.Print "Test des résultats pour le jour: " & testDate
    Debug.Print "-------------------------------"
    
    For Each rowData In result
        Debug.Print "Date Debut: " & rowData("Date Début")
        Debug.Print "Date Fin: " & rowData("Date Fin")
        Debug.Print "Heure Entree: " & rowData("Heure Entree")
        Debug.Print "Heure Sortie: " & rowData("Heure Sortie")
        Debug.Print "RR: " & rowData("RR")
        Debug.Print "Gain: " & rowData("Gain")
        Debug.Print "Key Trade: " & rowData("KeyTrade")
        Debug.Print "Commentary: " & rowData("Commentary")
        Debug.Print "Actif: " & rowData("Actif")
        Dim colIndex As Variant
        For Each colIndex In rowData.Keys
            If InStr(colIndex, "Screenshot") > 0 Then
                Debug.Print colIndex & ": " & rowData(colIndex)
            End If
        Next colIndex
        Debug.Print "-------------------------------"
    Next rowData
End Sub
