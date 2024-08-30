Attribute VB_Name = "PreparationDatasTest"
Sub TestPreparationDatas2()
    Dim filteredLines As New Collection
    Dim testData As Variant
    Dim resultDict As Scripting.Dictionary
    
    ' Ajouter des donn�es de test � filteredLines
    ' Les donn�es de test doivent �tre sous forme de tableau 2D, chaque ligne repr�sentant un trade
    ' Pour simplifier, nous utiliserons des valeurs al�atoires pour les donn�es de test
    For i = 1 To 100 ' Nombre de trades
        testData = Array("Date", "Heure", "Date", "Heure", Rnd() - 0.5) ' Chaque �l�ment de testData correspond � une colonne de votre feuille de calcul
        filteredLines.Add testData
    Next i
    
    ' Appeler la fonction de pr�paration des donn�es
    Set resultDict = preparationDatas(filteredLines)
    
    ' Afficher les r�sultats dans la fen�tre de l'�diteur VBA
    ' Vous pouvez remplacer MsgBox par Debug.Print pour afficher dans la fen�tre Output
    MsgBox "Nombre total de victoires : " & resultDict("WinRateDatas")("NbWin") & vbCrLf & _
           "Nombre total de d�faites : " & resultDict("WinRateDatas")("NbLoss")
    ' Afficher les donn�es pour les heures
    For Each key In resultDict("HoursData")
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("HoursData")(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & resultDict("HoursData")(key)("NbLoss")
    Next key
    
    ' Afficher les donn�es pour les jours
    For Each key In resultDict("DaysData")
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("DaysData")(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & resultDict("DaysData")(key)("NbLoss")
    Next key
    
    ' Afficher les donn�es pour les plages horaires
    For Each key In resultDict("TimeRangesData")
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("TimeRangesData")(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & resultDict("TimeRangesData")(key)("NbLoss")
    Next key
End Sub

Sub TestInitHoursData()
    Dim HoursData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des heures
    Set HoursData = InitHoursData()
    
    ' Afficher les donn�es initialis�es dans la fen�tre de l'�diteur VBA
    For Each key In HoursData
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & HoursData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & HoursData(key)("NbLoss")
    Next key
End Sub
'OK tester
Sub TestInitDaysData()
    Dim DaysData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des jours
    Set DaysData = InitDaysData()
    
    ' Afficher les donn�es initialis�es dans la fen�tre de l'�diteur VBA
    For Each key In DaysData
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & DaysData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & DaysData(key)("NbLoss")
    Next key
End Sub

Sub TestInitTimeRangesData()
    Dim TimeRangesData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des plages horaires
    Set TimeRangesData = InitTimeRangesData()
    
    ' Afficher les donn�es initialis�es dans la fen�tre de l'�diteur VBA
    For Each key In TimeRangesData
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & TimeRangesData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & TimeRangesData(key)("NbLoss")
    Next key
End Sub

'Fonction Tester
Sub TestCountWinsAndLosses()
    ' Cr�er une collection de lignes simul�es
    Dim filteredLines As New Collection
    Dim lineData As Variant
    Dim i As Integer
    Dim DDebut As Date
    Dim DFin As Date
    DDebut = dateValue("2023-01-01") ' Format: yyyy-mm-dd
    DFin = dateValue("2024-01-01") ' Format: yyyy-mm-dd
    Set filteredLines = FiltreRowByDate(DDebut, DFin, 1, "TrackRecord")
    
    ' Cr�er les dictionnaires n�cessaires
    Dim HoursData As Scripting.Dictionary
    Dim DaysData As Scripting.Dictionary
    Dim TimeRangesData As Scripting.Dictionary
    Dim tradesDatas As New Collection
    Dim WinRateDatas As New Scripting.Dictionary
    
    Set HoursData = InitHoursData()
    Set DaysData = InitDaysData()
    Set TimeRangesData = InitTimeRangesData()
    
    ' Appeler la fonction CountWinsAndLosses pour traiter les donn�es simul�es
    CountWinsAndLosses filteredLines, HoursData, DaysData, TimeRangesData, tradesDatas, WinRateDatas
    
    ' Afficher les r�sultats
    Dim key As Variant
    Debug.Print "Hours Data:"
    For Each key In HoursData.Keys
        Debug.Print key & ": NbWin = " & HoursData(key)("NbWin") & ", NbLoss = " & HoursData(key)("NbLoss")
    Next key
    
    Debug.Print "Days Data:"
    For Each key In DaysData.Keys
        Debug.Print key & ": NbWin = " & DaysData(key)("NbWin") & ", NbLoss = " & DaysData(key)("NbLoss")
    Next key
    
    Debug.Print "Time Ranges Data:"
    For Each key In TimeRangesData.Keys
        Debug.Print key & ": NbWin = " & TimeRangesData(key)("NbWin") & ", NbLoss = " & TimeRangesData(key)("NbLoss")
    Next key
    
    ' Afficher le nombre total de trades gagnants et perdants
    Debug.Print "Total Win Trades: " & tradesDatas.Count
End Sub

Sub TestPrintLineData()
    ' Cr�ez une collection pour tester
    Dim filteredLines As New Collection
    Dim lineData(1 To 1, 1 To 5) As Variant
    
    ' Remplir lineData avec des donn�es de test
    lineData(1, 1) = dateValue("2023-01-01")
    lineData(1, 2) = TimeValue("12:00:00")
    lineData(1, 3) = "TestData3"
    lineData(1, 4) = "TestData4"
    lineData(1, 5) = 1 ' Risk/Reward
    
    ' Ajouter lineData � filteredLines
    filteredLines.Add lineData
    
    ' Imprimer tous les �l�ments de lineData
    Dim i As Integer
    Dim j As Integer

    ' Assumer que lineData est un tableau 2D
    For i = LBound(lineData, 1) To UBound(lineData, 1)
        For j = LBound(lineData, 2) To UBound(lineData, 2)
            Debug.Print "lineData(" & i & ", " & j & ") = " & lineData(i, j)
        Next j
    Next i
End Sub

Sub TestPreparationDatas()
    Dim DateDebut As Date
    Dim DateFin As Date
    Dim dataDict As Scripting.Dictionary
    Dim key As Variant
    
    ' Sp�cifier les dates de d�but et de fin pour la p�riode de filtrage
    DateDebut = dateValue("2023-01-01") ' Format: yyyy-mm-dd
    DateFin = dateValue("2024-01-01") ' Format: yyyy-mm-dd
    
    ' Appeler la fonction preparationDatas pour obtenir le dictionnaire de donn�es
    Set dataDict = preparationDatas(DateDebut, DateFin)
    
    ' V�rifier si le dictionnaire de donn�es n'est pas vide
    If Not dataDict Is Nothing Then
        ' Imprimer le contenu de chaque dictionnaire inclus dans dataDict
        For Each key In dataDict.Keys
            Debug.Print "Contenu de " & key & ":"
            If TypeName(dataDict(key)) = "Dictionary" Then
                ' Si l'�l�ment est un dictionnaire, imprimer ses cl�s et valeurs
                PrintDictionary dataDict(key)
            ElseIf TypeName(dataDict(key)) = "Collection" Then
                ' Si l'�l�ment est une collection, imprimer ses �l�ments
                PrintCollection dataDict(key)
            Else
                ' Si l'�l�ment est d'un autre type, imprimer son contenu
                Debug.Print dataDict(key)
            End If
        Next key
    Else
        Debug.Print "Le dictionnaire de donn�es est vide."
    End If
End Sub

Sub PrintDictionary(dict As Scripting.Dictionary)
    Dim key As Variant
    
    ' Imprimer chaque cl� et sa valeur associ�e dans le dictionnaire
    For Each key In dict.Keys
        Debug.Print key & " NbWin: " & dict(key)("NbWin")
        Debug.Print key & " NbLoss: " & dict(key)("NbLoss")
    Next key
End Sub

Sub PrintCollection(coll As Collection)
    Dim item As Variant
    
    ' Imprimer chaque �l�ment dans la collection
    For Each item In coll
        Debug.Print item
    Next item
End Sub
