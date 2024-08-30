Attribute VB_Name = "TestPreparationDatas"
Sub TestPreparationDatas2()
    Dim filteredLines As New Collection
    Dim testData As Variant
    Dim resultDict As Scripting.Dictionary
    
    ' Ajouter des données de test à filteredLines
    ' Les données de test doivent être sous forme de tableau 2D, chaque ligne représentant un trade
    ' Pour simplifier, nous utiliserons des valeurs aléatoires pour les données de test
    For i = 1 To 100 ' Nombre de trades
        testData = Array("Date", "Heure", "Date", "Heure", Rnd() - 0.5) ' Chaque élément de testData correspond à une colonne de votre feuille de calcul
        filteredLines.Add testData
    Next i
    
    ' Appeler la fonction de préparation des données
    Set resultDict = preparationDatas(filteredLines)
    
    ' Afficher les résultats dans la fenêtre de l'éditeur VBA
    ' Vous pouvez remplacer MsgBox par Debug.Print pour afficher dans la fenêtre Output
    MsgBox "Nombre total de victoires : " & resultDict("WinRateDatas")("NbWin") & vbCrLf & _
           "Nombre total de défaites : " & resultDict("WinRateDatas")("NbLoss")
    ' Afficher les données pour les heures
    For Each key In resultDict("HoursData")
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("HoursData")(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & resultDict("HoursData")(key)("NbLoss")
    Next key
    
    ' Afficher les données pour les jours
    For Each key In resultDict("DaysData")
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("DaysData")(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & resultDict("DaysData")(key)("NbLoss")
    Next key
    
    ' Afficher les données pour les plages horaires
    For Each key In resultDict("TimeRangesData")
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & resultDict("TimeRangesData")(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & resultDict("TimeRangesData")(key)("NbLoss")
    Next key
End Sub

Sub TestInitHoursData()
    Dim HoursData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des heures
    Set HoursData = InitHoursData()
    
    ' Afficher les données initialisées dans la fenêtre de l'éditeur VBA
    For Each key In HoursData
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & HoursData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & HoursData(key)("NbLoss")
    Next key
End Sub

Sub TestInitDaysData()
    Dim DaysData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des jours
    Set DaysData = InitDaysData()
    
    ' Afficher les données initialisées dans la fenêtre de l'éditeur VBA
    For Each key In DaysData
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & DaysData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & DaysData(key)("NbLoss")
    Next key
End Sub

Sub TestInitTimeRangesData()
    Dim TimeRangesData As Scripting.Dictionary
    Dim key As Variant
    
    ' Appeler la fonction d'initialisation des plages horaires
    Set TimeRangesData = InitTimeRangesData()
    
    ' Afficher les données initialisées dans la fenêtre de l'éditeur VBA
    For Each key In TimeRangesData
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & TimeRangesData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & TimeRangesData(key)("NbLoss")
    Next key
End Sub

Sub TestCountWinsAndLosses()
    Dim filteredLines As New Collection
    Dim testData As Variant
    Dim HoursData As Scripting.Dictionary
    Dim DaysData As Scripting.Dictionary
    Dim TimeRangesData As Scripting.Dictionary
    Dim tradesDatas As New Collection
    Dim WinRateDatas As Scripting.Dictionary
    
    ' Ajouter des données de test à filteredLines
    ' Les données de test doivent être sous forme de tableau 2D, chaque ligne représentant un trade
    ' Pour simplifier, nous utiliserons des valeurs aléatoires pour les données de test
    For i = 1 To 100 ' Nombre de trades
        testData = Array("Date", "Heure", "Date", "Heure", Rnd() - 0.5) ' Chaque élément de testData correspond à une colonne de votre feuille de calcul
        filteredLines.Add testData
    Next i
    
    ' Initialiser les dictionnaires pour les heures, les jours et les plages horaires
    Set HoursData = InitHoursData()
    Set DaysData = InitDaysData()
    Set TimeRangesData = InitTimeRangesData()
    
    ' Appeler la fonction de comptage des victoires et des défaites
    CountWinsAndLosses filteredLines, HoursData, DaysData, TimeRangesData, tradesDatas, WinRateDatas
    
    ' Afficher les données comptées dans des boîtes de message
    ' Vous pouvez remplacer MsgBox par Debug.Print pour afficher dans la fenêtre Output
    ' Afficher les données pour les heures
    For Each key In HoursData
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & HoursData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & HoursData(key)("NbLoss")
    Next key
    
    ' Afficher les données pour les jours
    For Each key In DaysData
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & DaysData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & DaysData(key)("NbLoss")
    Next key
    
    ' Afficher les données pour les plages horaires
    For Each key In TimeRangesData
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & TimeRangesData(key)("NbWin") & vbCrLf & _
               "Nombre de défaites : " & TimeRangesData(key)("NbLoss")
    Next key
End Sub





