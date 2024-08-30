Attribute VB_Name = "TrieEtPreparationTest"
Sub TestFiltreRowByDate()
    Dim DateDebut As Date
    Dim DateFin As Date
    Dim ColumnDate As Integer
    Dim WorksheetName As String
    Dim filteredLines As Collection
    Dim lineData As Variant
    Dim i As Integer
    Dim j As Integer
    
    ' Définir les paramètres de test
    DateDebut = DateSerial(2023, 1, 1) ' Début de la plage de dates
    DateFin = DateSerial(2023, 12, 31) ' Fin de la plage de dates
    ColumnDate = 1 ' Colonne de la date dans la feuille de calcul (A = 1)
    WorksheetName = "TrackRecord" ' Nom de la feuille de calcul à tester

    ' Appeler la fonction FiltreRowByDate
    Set filteredLines = FiltreRowByDate(DateDebut, DateFin, ColumnDate, WorksheetName)
    
    ' Vérifier si la collection retournée n'est pas vide
    If Not filteredLines Is Nothing Then
        If filteredLines.Count > 0 Then
            Debug.Print "Lignes filtrées :"
            ' Parcourir chaque ligne filtrée et imprimer les valeurs
            For Each lineData In filteredLines
                For j = LBound(lineData, 2) To UBound(lineData, 2)
                    Debug.Print lineData(1, j); " ";
                Next j
                Debug.Print ' Nouvelle ligne pour chaque ligne de données
            Next lineData
        Else
            Debug.Print "Aucune ligne trouvée pour la plage de dates spécifiée."
        End If
    Else
        Debug.Print "La collection retournée est Nothing."
    End If
End Sub


Sub TestPreparationDatas()
    Dim filteredLines As Collection
    Dim dataDict As Scripting.Dictionary
    Dim WinRateDatas As Scripting.Dictionary
    Dim NbWin As Integer
    Dim nbLoss As Integer
    Dim nbTie As Integer
    Dim tradesDatas As Collection

    ' Appeler la fonction FiltreDatas pour obtenir les lignes filtrées
    Set filteredLines = FiltreDatas()
    
    ' Vérifier que des lignes ont été filtrées
    If filteredLines Is Nothing Then
        Debug.Print "Aucune ligne filtrée."
        Exit Sub
    End If
    
    ' Appeler la fonction PreparationDatas pour traiter les lignes filtrées
    Set dataDict = preparationDatas(filteredLines)
    
    ' Vérifier que le dictionnaire de données a été rempli correctement
    If dataDict Is Nothing Then
        Debug.Print "Aucune donnée extraite."
        Exit Sub
    End If
    
    ' Récupérer les données de win/loss/tie
    Set WinRateDatas = dataDict("WinRateDatas")
    NbWin = WinRateDatas("NbWin")
    nbLoss = WinRateDatas("NbLoss")
    
    ' Afficher les résultats dans la fenêtre d'exécution
    Debug.Print "NbWin: " & NbWin
    Debug.Print "NbLoss: " & nbLoss

    Debug.Print "Tout roule ma poule"
    
    If dataDict.Exists("TradesDatas") Then
        ' Récupérer la collection TradesData
        Set tradesData = dataDict("TradesDatas")
        
        ' Afficher le contenu de TradesData
        Debug.Print "Contenu de TradesData :"
        For Each lineData In tradesData
            Debug.Print lineData
        Next lineData
    End If
End Sub

