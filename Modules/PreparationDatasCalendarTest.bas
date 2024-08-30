Attribute VB_Name = "PreparationDatasCalendarTest"
Public CalendarLabelDictionary As Object

Sub InitializeCalendarLabelDictionary()
    Set CalendarLabelDictionary = CreateObject("Scripting.Dictionary")
End Sub


Sub TestInitDictDayOfMonth()
    Dim year As Integer
    Dim month As Integer
    Dim dict As Scripting.Dictionary
    Dim i As Integer

    year = 2024
    month = 6 ' Juin

    ' Initialiser le dictionnaire
    Set dict = InitDictDayOfMonth(year, month)
    
    ' Afficher le contenu du dictionnaire
    Debug.Print "Dictionnaire pour " & month & "/" & year & ":"
    
    For i = 1 To dict.Count
        Dim subDict As Scripting.Dictionary
        Set subDict = dict(i)
        
        Debug.Print "Jour " & i & ":"
        Debug.Print "  nbwin: " & subDict("nbwin")
        Debug.Print "  nbloose: " & subDict("nbloose")
        Debug.Print "  RR: " & subDict("RR")
        Debug.Print "  Trades (nombre d'éléments): " & subDict("Trades").Count
    Next i
End Sub


Sub TestInsertDatasInDictionnary()
    ' Initialiser les variables
    Dim dict As Scripting.Dictionary
    Dim filteredRows As Collection
    Dim i As Integer
    Dim subDict As Scripting.Dictionary
    Dim dictInit As Scripting.Dictionary

    ' Initialiser le dictionnaire pour le mois de mai 2024
    Set dictInit = InitDictDayOfMonth(2024, 5)
    
    ' Filtrer les lignes par date pour le mois de mai 2024 (supposons que les dates sont dans la colonne A)
    Set filteredRows = FiltreRowByDate(DateSerial(2024, 5, 1), DateSerial(2024, 5, 31), 1, "Feuille1")
    
    ' Insérer les données dans le dictionnaire
    InsertDatasInDictionnary dictInit, filteredRows
    
    ' Afficher le contenu du dictionnaire pour vérification
    Debug.Print "Contenu du dictionnaire après l'insertion des données filtrées :"
    For i = 1 To dictInit.Count
        Set subDict = dictInit(i)
        Debug.Print "Jour " & i & ":"
        Debug.Print "  nbwin = " & subDict("nbwin")
        Debug.Print "  nbloose = " & subDict("nbloose")
        Debug.Print "  RR = " & subDict("RR")
        Debug.Print "  Trades (nombre d'éléments) = " & subDict("Trades").Count
    Next i
End Sub


Sub TestPreparationDatasBtnDay()
    ' Initialiser les variables
    Dim year As Integer
    Dim month As Integer
    Dim dict As Scripting.Dictionary
    Dim i As Integer
    Dim subDict As Scripting.Dictionary
    
    ' Spécifier l'année et le mois pour le test (mai 2024)
    year = 2024
    month = 5
    
    ' Appeler la fonction pour préparer les données et obtenir le dictionnaire initialisé
    Set dict = PreparationDatasBtnDay(year, month)
    
    ' Vérifier si le dictionnaire a été correctement initialisé
    If Not dict Is Nothing Then
        ' Afficher le contenu du dictionnaire pour vérification
        Debug.Print "Contenu du dictionnaire après la préparation des données :"
        For i = 1 To dict.Count
            Set subDict = dict(i)
            Debug.Print "Jour " & i & ":"
            Debug.Print "  nbwin = " & subDict("nbwin")
            Debug.Print "  nbloose = " & subDict("nbloose")
            Debug.Print "  RR = " & subDict("RR")
            Debug.Print "  Trades (nombre d'éléments) = " & subDict("Trades").Count
            
            Dim trade As Variant
            For Each trade In subDict("Trades")
                Debug.Print "    Trade RR Value: " & trade
            Next trade
        Next i
    Else
        ' Afficher un message d'erreur si le dictionnaire est indéfini
        Debug.Print "Erreur : Le dictionnaire n'a pas été correctement initialisé."
    End If
End Sub
