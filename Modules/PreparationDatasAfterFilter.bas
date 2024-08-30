Attribute VB_Name = "PreparationDatasAfterFilter"
'---------------------------------MainFunction------------------------------

Function preparationDatas(DateDebut As Date, DateFin As Date) As Scripting.Dictionary
    Dim dataDict As New Scripting.Dictionary
    Dim tradesDatas As New Collection
    Dim HoursData As Scripting.Dictionary
    Dim DaysData As Scripting.Dictionary
    Dim GeneralData As Scripting.Dictionary
    Dim SessionsData As Scripting.Dictionary
    Dim filteredLines As Collection
    ' Initialiser les dictionnaires pour les heures et les jours
    Set HoursData = InitHoursData()
    Set DaysData = InitDaysData()
    Set SessionsData = InitTimeRangesData()
    Set GeneralData = InitGeneralData()
    'filtrer les lines des dates
    Set filteredLines = FiltreRowByDate(DateDebut, DateFin, 1, "TrackRecord")
    
    ' Compter les victoires et les défaites pour chaque heure et chaque jour
    CountWinsAndLosses filteredLines, GeneralData, HoursData, DaysData, SessionsData, tradesDatas
    
 
    ' Ajouter les dictionnaires et les collections à la structure de données principale
    dataDict.Add "WinRateDatas", GeneralData
    dataDict.Add "TradesDatas", tradesDatas
    dataDict.Add "HoursData", HoursData
    dataDict.Add "DaysData", DaysData
    dataDict.Add "SessionsData", SessionsData
    ' Retourner la collection de données extraites
    Set preparationDatas = dataDict
End Function

Function CountWinsAndLosses(filteredLines As Collection, ByRef GeneralData As Scripting.Dictionary, ByRef HoursData As Scripting.Dictionary, ByRef DaysData As Scripting.Dictionary, ByRef TimeRangesData As Scripting.Dictionary, ByRef tradesDatas As Collection) As Integer

    ' Définition des différentes colonnes de Trackrecord
    Dim dateDebutTradeColonne As Integer
    Dim heureDebutTradeColonne As Integer
    Dim riskRewardColonne As Integer
    dateDebutTradeColonne = 1
    heureDebutTradeColonne = 2
    riskRewardColonne = 5
    
    ' Variables Win/Loss
    Dim NbWin As Integer
    Dim nbLoss As Integer
    NbWin = 0
    nbLoss = 0
    
    ' Boucle pour compter les wins, losses et remplir les données des heures, des jours et des plages horaires
    For Each lineData In filteredLines
        Dim tradeHour As String
        Dim tradeDay As String
        Dim tradeTimeRange As String
        
        tradeHour = Format(lineData(1, heureDebutTradeColonne), "hh") & ":00"
        tradeDay = WeekdayName(Weekday(lineData(1, dateDebutTradeColonne), vbMonday))
    
        ' Identifier la plage horaire correspondante
        If lineData(1, heureDebutTradeColonne) >= TimeValue("08:00:00") And lineData(1, heureDebutTradeColonne) < TimeValue("14:00:00") Then
            tradeTimeRange = "8:00-14:00"
        ElseIf lineData(1, heureDebutTradeColonne) >= TimeValue("14:00:00") And lineData(1, heureDebutTradeColonne) < TimeValue("21:00:00") Then
            tradeTimeRange = "14:00-21:00"
        Else
            tradeTimeRange = "21:00-8:00"
        End If
        
        If lineData(1, riskRewardColonne) > 0 Then
            NbWin = NbWin + 1
        ElseIf lineData(1, riskRewardColonne) < 0 Then
            nbLoss = nbLoss + 1
        End If
        
        ' Mise à jour des données pour les heures, les jours et les plages horaires
        Dim trade As Integer
        trade = lineData(1, riskRewardColonne)
        
        Call UpdateRRDico(HoursData, tradeHour, trade)
        Call UpdateRRDico(DaysData, tradeDay, trade)
        Call UpdateRRDico(TimeRangesData, tradeTimeRange, trade)
        
        'Mise à jour données GénéralData
        Call UpdateGeneralDataRRDico(GeneralData, trade)
    Next lineData

    CountWinsAndLosses = NbWin + nbLoss ' Return the total number of trades
End Function


'--------------------------INITIALISATION DICTIONNARY--------------------



Function InitGeneralData() As Scripting.Dictionary
    Dim GeneralData As Scripting.Dictionary
    Set GeneralData = New Scripting.Dictionary
    Call InitAddDict(GeneralData)
    Set InitGeneralData = GeneralData
End Function

Function InitHoursData() As Scripting.Dictionary
    Dim HoursData As New Scripting.Dictionary
    Dim hour As Integer
    
    ' Initialiser les dictionnaires pour les heures
    For hour = 0 To 23
        Dim hourDict As Scripting.Dictionary
        Set hourDict = New Scripting.Dictionary
        Call InitAddDict(hourDict)
        ' Utilisation du formatage "hh:00" pour les clés du dictionnaire
        HoursData.Add Format(hour, "00") & ":00", hourDict
    Next hour
    
    Set InitHoursData = HoursData
End Function

'DayDatas:{"Lundi":{NbWin:...,NbLoose:...,Trade:[},"Mardi:{....}
Function InitDaysData() As Scripting.Dictionary
    Dim DaysData As New Scripting.Dictionary
    Dim day As Variant
    Dim daysOfWeek As Variant
    
    'Voir si je peux rendre daysOfWeek plus responsive notamment en utilisant une fonction VBA
    daysOfWeek = Array("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    
    ' Initialiser les dictionnaires pour les jours de la semaine
    For Each day In daysOfWeek
        Dim dayDict As Scripting.Dictionary
        Set dayDict = New Scripting.Dictionary
        Call InitAddDict(dayDict)
        DaysData.Add day, dayDict
    Next day
    
    Set InitDaysData = DaysData
End Function

'Tester
 'TimeRangesDatas:{"8:00-"14:00":{NbWin:...,NbLoose:...,Trades[]}
Function InitTimeRangesData() As Scripting.Dictionary
    Dim TimeRangesData As New Scripting.Dictionary
    
    ' Initialiser les dictionnaires pour les plages horaires
    Dim range As Variant
    For Each range In Array("8:00-14:00", "14:00-21:00", "21:00-8:00")
        Dim rangeDict As Scripting.Dictionary
        Set rangeDict = New Scripting.Dictionary
        Call InitAddDict(rangeDict)
        TimeRangesData.Add range, rangeDict
    Next range
    
    Set InitTimeRangesData = TimeRangesData
End Function

Function InitAddDict(ByRef dict As Scripting.Dictionary)
        dict.Add "TotalGain", 0
        dict.Add "TotalLoss", 0
        dict.Add "NbWin", 0
        dict.Add "NbLoss", 0
        dict.Add "Trades", New Collection
        dict.Add "TotalRRGain", 0
        dict.Add "TotalRRLosses", 0
        dict.Add "TotalMonneyGain", 0
        dict.Add "TotalMonneyLosses", 0
End Function




'------------------------------- UPDATE DICO FUNCTIONS--------------------------------------------



Function UpdateRRDico(ByRef dataDict As Scripting.Dictionary, key As String, value As Variant)
    
    Call UpdateDataBase(dataDict, key, "NbWin", IIf(value > 0, 1, 0))
    Call UpdateDataBase(dataDict, key, "NbLoss", IIf(value < 0, 1, 0))
    Call UpdateDataBase(dataDict, key, "Trades", value)
    Call UpdateDataBase(dataDict, key, "TotalRRGain", IIf(value > 0, value, 0))
    Call UpdateDataBase(dataDict, key, "TotalRRLosses", IIf(value < 0, value, 0))
   
    
    
    
End Function

Function UpdateMoneyDico(ByRef dataDict As Scripting.Dictionary, key As String, value As Variant)
    
    UpdateDataBase dataDict, key, "TotalMonneyGain", IIf(value > 0, value, 0)
    UpdateDataBase dataDict, key, "TotalMonneyLosses", IIf(value < 0, value, 0)
    
    
End Function




Function UpdateDataBase(ByRef dataDict As Scripting.Dictionary, key As String, subKey As String, value As Variant)
    If Not dataDict.Exists(key) Then
       Debug.Print ("La clé n'existe pas")
        Exit Function
    End If
    
    If Not dataDict(key).Exists(subKey) Then
        Debug.Print ("La clé n'existe pas")
        Exit Function
    End If
    
    If subKey = "Trades" Then
        dataDict(key)(subKey).Add value
    Else
        dataDict(key)(subKey) = dataDict(key)(subKey) + value
    End If
End Function
            'Update juste du dico GeneralDatas
Function UpdateGeneralDataRRDico(ByRef dataDict As Scripting.Dictionary, value As Variant)
    Call UpdateGeneralData(dataDict, "NbWin", IIf(value > 0, 1, 0))
    Call UpdateGeneralData(dataDict, "NbLoss", IIf(value < 0, 1, 0))
    Call UpdateGeneralData(dataDict, "Trades", value)
    Call UpdateGeneralData(dataDict, "TotalRRGain", IIf(value > 0, value, 0))
    Call UpdateGeneralData(dataDict, "TotalRRLosses", IIf(value < 0, value, 0))
   
End Function

Function UpdateGeneralMoneyDico(ByRef dataDict As Scripting.Dictionary, value As Variant)
    
    UpdateDataBase dataDict, "TotalMonneyGain", IIf(value > 0, value, 0)
    UpdateDataBase dataDict, "TotalMonneyLosses", IIf(value < 0, value, 0)
    
    
End Function
Function UpdateGeneralData(ByRef dataDict As Scripting.Dictionary, key As String, value As Variant)
    If Not dataDict.Exists(key) Then
        Debug.Print "La clé n'existe pas"
        Exit Function
    End If
    
    If IsObject(dataDict(key)) And TypeName(dataDict(key)) = "Collection" Then
        dataDict(key).Add value
    Else
        dataDict(key) = dataDict(key) + value
    End If
End Function
