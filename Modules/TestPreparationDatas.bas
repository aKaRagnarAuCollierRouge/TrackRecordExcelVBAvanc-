Attribute VB_Name = "TestPreparationDatas"
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

Sub TestCountWinsAndLosses()
    Dim filteredLines As New Collection
    Dim testData As Variant
    Dim HoursData As Scripting.Dictionary
    Dim DaysData As Scripting.Dictionary
    Dim TimeRangesData As Scripting.Dictionary
    Dim tradesDatas As New Collection
    Dim WinRateDatas As Scripting.Dictionary
    
    ' Ajouter des donn�es de test � filteredLines
    ' Les donn�es de test doivent �tre sous forme de tableau 2D, chaque ligne repr�sentant un trade
    ' Pour simplifier, nous utiliserons des valeurs al�atoires pour les donn�es de test
    For i = 1 To 100 ' Nombre de trades
        testData = Array("Date", "Heure", "Date", "Heure", Rnd() - 0.5) ' Chaque �l�ment de testData correspond � une colonne de votre feuille de calcul
        filteredLines.Add testData
    Next i
    
    ' Initialiser les dictionnaires pour les heures, les jours et les plages horaires
    Set HoursData = InitHoursData()
    Set DaysData = InitDaysData()
    Set TimeRangesData = InitTimeRangesData()
    
    ' Appeler la fonction de comptage des victoires et des d�faites
    CountWinsAndLosses filteredLines, HoursData, DaysData, TimeRangesData, tradesDatas, WinRateDatas
    
    ' Afficher les donn�es compt�es dans des bo�tes de message
    ' Vous pouvez remplacer MsgBox par Debug.Print pour afficher dans la fen�tre Output
    ' Afficher les donn�es pour les heures
    For Each key In HoursData
        MsgBox "Heure : " & key & vbCrLf & _
               "Nombre de victoires : " & HoursData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & HoursData(key)("NbLoss")
    Next key
    
    ' Afficher les donn�es pour les jours
    For Each key In DaysData
        MsgBox "Jour : " & key & vbCrLf & _
               "Nombre de victoires : " & DaysData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & DaysData(key)("NbLoss")
    Next key
    
    ' Afficher les donn�es pour les plages horaires
    For Each key In TimeRangesData
        MsgBox "Plage horaire : " & key & vbCrLf & _
               "Nombre de victoires : " & TimeRangesData(key)("NbWin") & vbCrLf & _
               "Nombre de d�faites : " & TimeRangesData(key)("NbLoss")
    Next key
End Sub





