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
    
    ' D�finir les param�tres de test
    DateDebut = DateSerial(2023, 1, 1) ' D�but de la plage de dates
    DateFin = DateSerial(2023, 12, 31) ' Fin de la plage de dates
    ColumnDate = 1 ' Colonne de la date dans la feuille de calcul (A = 1)
    WorksheetName = "TrackRecord" ' Nom de la feuille de calcul � tester

    ' Appeler la fonction FiltreRowByDate
    Set filteredLines = FiltreRowByDate(DateDebut, DateFin, ColumnDate, WorksheetName)
    
    ' V�rifier si la collection retourn�e n'est pas vide
    If Not filteredLines Is Nothing Then
        If filteredLines.Count > 0 Then
            Debug.Print "Lignes filtr�es :"
            ' Parcourir chaque ligne filtr�e et imprimer les valeurs
            For Each lineData In filteredLines
                For j = LBound(lineData, 2) To UBound(lineData, 2)
                    Debug.Print lineData(1, j); " ";
                Next j
                Debug.Print ' Nouvelle ligne pour chaque ligne de donn�es
            Next lineData
        Else
            Debug.Print "Aucune ligne trouv�e pour la plage de dates sp�cifi�e."
        End If
    Else
        Debug.Print "La collection retourn�e est Nothing."
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

    ' Appeler la fonction FiltreDatas pour obtenir les lignes filtr�es
    Set filteredLines = FiltreDatas()
    
    ' V�rifier que des lignes ont �t� filtr�es
    If filteredLines Is Nothing Then
        Debug.Print "Aucune ligne filtr�e."
        Exit Sub
    End If
    
    ' Appeler la fonction PreparationDatas pour traiter les lignes filtr�es
    Set dataDict = preparationDatas(filteredLines)
    
    ' V�rifier que le dictionnaire de donn�es a �t� rempli correctement
    If dataDict Is Nothing Then
        Debug.Print "Aucune donn�e extraite."
        Exit Sub
    End If
    
    ' R�cup�rer les donn�es de win/loss/tie
    Set WinRateDatas = dataDict("WinRateDatas")
    NbWin = WinRateDatas("NbWin")
    nbLoss = WinRateDatas("NbLoss")
    
    ' Afficher les r�sultats dans la fen�tre d'ex�cution
    Debug.Print "NbWin: " & NbWin
    Debug.Print "NbLoss: " & nbLoss

    Debug.Print "Tout roule ma poule"
    
    If dataDict.Exists("TradesDatas") Then
        ' R�cup�rer la collection TradesData
        Set tradesData = dataDict("TradesDatas")
        
        ' Afficher le contenu de TradesData
        Debug.Print "Contenu de TradesData :"
        For Each lineData In tradesData
            Debug.Print lineData
        Next lineData
    End If
End Sub

