Attribute VB_Name = "PreparationDatasCalendar"
'OKAY Tester
Function InitDictDayOfMonth(year As Integer, month As Integer) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim daysInMonth As Integer
    daysInMonth = day(DateSerial(year, month + 1, 0)) ' Calcul du nombre de jours dans le mois

    Dim i As Integer
    For i = 1 To daysInMonth
        Dim subDict As Scripting.Dictionary
        Set subDict = New Scripting.Dictionary
        subDict.Add "nbwin", 0
        subDict.Add "nbloose", 0
        subDict.Add "RR", 0
        
        ' Ajouter une collection vide pour les Trades
        Dim trades As Collection
        Set trades = New Collection
        subDict.Add "Trades", trades
        
        dict.Add i, subDict
    Next i

    Set InitDictDayOfMonth = dict
End Function

Function InsertDatasInDictionnary(ByVal dictInit As Scripting.Dictionary, ByVal FiltredRows As Collection)
    Dim RRColonne As Integer
    Dim ColonneDate As Integer
    Dim row As Variant
    Dim dateValue As Date
    Dim dayOfMonth As Integer
    Dim rrValue As Double
    Dim item As Variant
    Dim i As Integer
    Dim wsTrackrecord As Worksheet
    Set wsTrackrecord = ThisWorkbook.Worksheets("Trackrecord")
    RRColonne = FindIndiceCollumnWithTable(wsTrackrecord, "Tableau1", "RR")
    ColonneDate = FindIndiceCollumnWithTable(wsTrackrecord, "Tableau1", "Date Début")
    
    For Each row In FiltredRows
      
        
        dateValue = row(1, ColonneDate)
      
        
        dayOfMonth = day(dateValue)
      
        
        rrValue = row(1, RRColonne)
        
        
        ' Mettre à jour le dictionnaire
        If Not dictInit.Exists(dayOfMonth) Then
            Set dictInit(dayOfMonth) = CreateObject("Scripting.Dictionary")
            dictInit(dayOfMonth).Add "nbwin", 0
            dictInit(dayOfMonth).Add "nbloose", 0
            dictInit(dayOfMonth).Add "RR", 0
            Set dictInit(dayOfMonth)("Trades") = New Collection
        End If
        
        dictInit(dayOfMonth)("nbwin") = dictInit(dayOfMonth)("nbwin") + 1
        dictInit(dayOfMonth)("nbloose") = dictInit(dayOfMonth)("nbloose") + 1
        dictInit(dayOfMonth)("RR") = dictInit(dayOfMonth)("RR") + rrValue
        dictInit(dayOfMonth)("Trades").Add rrValue
    Next row
End Function





' Mettre cette fonction dans le calendar??? Sinon double échange entre 2 fichiers pas une bonne pratique
Function PreparationDatasBtnDay(ByVal year As Integer, ByVal month As Integer) As Scripting.Dictionary
    Dim wsTrackrecord As Worksheet
    Dim RowFiltred As Collection
    Dim FirstMonthDayDate As Date
    Dim LastMonthDayDate As Date
    Dim dates As Variant
    
    Set Calendar.dictDatasDay = InitDictDayOfMonth(year, month)
    dates = GetFirstAndLastDayOfMonth(year, month)
    Set wsTrackrecord = ThisWorkbook.Worksheets("TrackRecord")
    FirstMonthDayDate = dates(0)
    LastMonthDayDate = dates(1)
    Set RowFiltred = FiltreRowByDate(FirstMonthDayDate, LastMonthDayDate, 1, "TrackRecord")
    Call InsertDatasInDictionnary(Calendar.dictDatasDay, RowFiltred)
    Set PreparationDatasBtnDay = Calendar.dictDatasDay

End Function
