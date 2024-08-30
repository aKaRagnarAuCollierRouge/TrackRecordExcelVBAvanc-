Attribute VB_Name = "DBFetchTest"
Sub TestLoadJSONObject()
    Dim yearExisting As Integer
    Dim yearNonExisting As Integer
    Dim jsonObject As Object
    Dim fso As Object
    Dim jsonfile As String
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Année avec un fichier JSON existant
    yearExisting = 2024
    jsonfile = ThisWorkbook.Path & "\DB\" & CStr(yearExisting) & ".json"
    Debug.Print (jsonfile)
    ' Vérifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        Debug.Print "Fichier JSON existant trouvé pour l'année " & yearExisting
        Set jsonObject = Loadjson(yearExisting)
        If Not jsonObject Is Nothing Then
            Debug.Print "Test avec fichier JSON existant (année " & yearExisting & "):"
            ' Afficher le contenu JSON dans la fenêtre de débogage
            PrintJSONObject jsonObject
        Else
            Debug.Print "Test avec fichier JSON existant (année " & yearExisting & "): Chargement JSON échoué."
        End If
    Else
        Debug.Print "Fichier JSON non trouvé pour l'année " & yearExisting
    End If
    
    ' Année avec un fichier JSON non existant
    yearNonExisting = 2025
    jsonfile = ThisWorkbook.Path & "\DB\" & CStr(yearNonExisting) & ".json"
    
    ' Vérifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        Debug.Print "Fichier JSON existant trouvé pour l'année " & yearNonExisting
        Set jsonObject = Loadjson(yearNonExisting)
        If Not jsonObject Is Nothing Then
            Debug.Print "Test avec fichier JSON non existant (année " & yearNonExisting & "):"
            ' Afficher le contenu JSON dans la fenêtre de débogage
            PrintJSONObject jsonObject
        Else
            Debug.Print "Test avec fichier JSON non existant (année " & yearNonExisting & "): Chargement JSON échoué."
        End If
    Else
        Debug.Print "Fichier JSON non trouvé pour l'année " & yearNonExisting
    End If
End Sub

' Fonction pour afficher le contenu JSON dans la fenêtre de débogage
Sub PrintJSONObject(jsonObject As Object)
    Dim key As Variant
    
    For Each key In jsonObject.Keys
        ' Afficher chaque clé et sa valeur dans la fenêtre de débogage
        Debug.Print key & ": " & jsonObject(key)
    Next key
End Sub



Sub TestFolderExists()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print (GlobalsVariables.folderPathCommentary)
    
    ' Vérifier si le chemin existe
    If fso.FolderExists(GlobalsVariables.folderPathCommentary) Then
        Debug.Print ("Le chemin existe")
    Else
        Debug.Print ("Le chemin n'existe pas")
    End If
  
End Sub

Sub TestFileExists()
    Dim fso As Object
    Dim FilePath As String
    
    ' Initialiser l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Chemin du fichier à vérifier
    FilePath = GlobalsVariables.folderPathCommentary + "\2024.json"
    Debug.Print (FilePath)
    ' Vérifier si le fichier existe
    If fso.FileExists(FilePath) Then
        Debug.Print "Le fichier existe."
    Else
        Debug.Print "Le fichier n'existe pas."
    End If
End Sub

Sub TestGetKeyTradesDay()
    Dim folderPath As String
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim keyTrades As Collection
    Dim jsonObject As Object
    Dim jsonfile As String
    Dim i As Integer
    
    ' Définir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Définir l'année, le mois et le jour de test
    year = 2024
    month = 1
    day = 1
    
    ' Charger le fichier JSON de l'année spécifiée
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et récupérer les KeyTrades
    Set keyTrades = GetKeyTradesDay(jsonObject, month, day)
    
    ' Afficher les KeyTrades dans une boîte de message
    Dim keyTradeList As String
    keyTradeList = "KeyTrades : " & vbCrLf
    For i = 1 To keyTrades.Count
        keyTradeList = keyTradeList & keyTrades(i) & vbCrLf
    Next i
    
    MsgBox keyTradeList, vbInformation, "Test GetKeyTradesDay"
End Sub

Sub TestGetCommentaryDay()
    Dim folderPath As String
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim Commentary As String
    Dim jsonObject As Object
    Dim jsonfile As String
    
    ' Définir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Définir l'année, le mois et le jour de test
    year = 2024
    month = 1
    day = 1
    
    ' Charger le fichier JSON de l'année spécifiée
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et récupérer le commentaire du jour
    Commentary = GetCommentaryDay(jsonObject, month, day)
    
    ' Afficher le commentaire dans une boîte de message
    MsgBox "Commentaire : " & vbCrLf & Commentary, vbInformation, "Test GetCommentaryDay"
End Sub

Sub TestFindCommentMonth()
    Dim folderPath As String
    Dim year As Integer
    Dim month As Integer
    Dim result As String
    Dim jsonObject As Object
    Dim jsonfile As String
    
    ' Définir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Définir l'année et le mois de test
    year = 2024
    month = 1
    
    ' Charger le fichier JSON de l'année spécifiée
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et récupérer le résultat
    result = FindCommentMonth(jsonObject, month)
    
    ' Afficher le résultat dans une boîte de message
    MsgBox result, vbInformation, "Test FindCommentByYearMonth"
End Sub

Sub TestFindCommentYear()
    Dim folderPath As String
    Dim year As Integer
    Dim result As String
    Dim jsonObject As Object
    Dim jsonfile As String
    
    ' Définir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Définir l'année de test
    year = 2024
    
    ' Charger le fichier JSON de l'année spécifiée
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et récupérer le résultat
    result = FindCommentYear(jsonObject)
    
    ' Afficher le résultat dans une boîte de message
    MsgBox result, vbInformation, "Test FindCommentByYear"
End Sub



Option Explicit

' Déclaration de la variable globale pour le chemin du dossier
Public folderPath As String

' Cette procédure sera exécutée à l'ouverture du classeur
Private Sub Workbook_Open()
    ' Initialiser la variable globale avec le chemin du dossier contenant ce classeur
    folderPath = ThisWorkbook.Path
End Sub

' Fonction pour lire le contenu d'un fichier
Function LireFichier(FilePath As String) As String
    Dim fileContent As String
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open FilePath For Input As fileNumber
    fileContent = Input$(LOF(fileNumber), fileNumber)
    Close fileNumber
    LireFichier = fileContent
End Function

' Fonction pour charger un objet JSON à partir d'un fichier pour une année spécifique
Function LoadjsonObject(year As Integer) As Object
    Dim jsonfile As String
    Dim fso As Object
    Dim fileStream As Object
    Dim jsonObject As Object
    Dim FileExists As Variant

    ' Définir le chemin du fichier JSON de l'année spécifiée
    jsonfile = folderPath & "\" & year & ".json"
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Vérifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        ' Lire le fichier JSON existant
        Set fileStream = fso.OpenTextFile(jsonfile, 1) ' 1 = ForReading
        Set jsonObject = JsonConverter.ParseJson(fileStream.ReadAll)
        fileStream.Close
        
        ' Indiquer que le fichier JSON a été chargé avec succès
        Set LoadjsonObject = jsonObject
    Else
        ' Indiquer que le fichier JSON n'existe pas
        Set LoadjsonObject = Nothing
    End If
End Function

' Fonction pour récupérer les données pour un mois et un jour spécifiques à partir d'un objet JSON
Function GetDictionnaryDay(jsonObject As Object, month As Integer, day As Integer) As Dictionary
    Dim result As Dictionary
    Set result = New Dictionary
    
    ' Vérifier si le mois existe dans l'objet JSON
    If jsonObject.Exists(CStr(month)) Then
        ' Vérifier si le jour existe dans l'objet JSON pour le mois spécifié
        If jsonObject(CStr(month)).Exists(CStr(day)) Then
            Set result = jsonObject(CStr(month))(CStr(day))
        End If
    End If
    
    Set GetDictionnaryDay = result
End Function

' Fonction de test pour GetDictionnaryDay
Sub TestGetDictionnaryDay()
    Dim year As Integer
    Dim month As Integer
    Dim day As Integer
    Dim jsonObject As Object
    Dim resultat As Dictionary
    Dim jsonContent As String
    Dim jsonfile As String
    Dim keyTrade As Variant
    Dim keyTradeStr As String
    
    ' Spécifiez l'année, le mois et le jour que vous voulez tester
    year = 2024
    month = 1
    day = 1
    
    ' Charger l'objet JSON pour l'année spécifiée
    Set jsonObject = LoadjsonObject(year)
    
    ' Vérifier si l'objet JSON a été chargé
    If Not jsonObject Is Nothing Then
        ' Appeler la fonction GetDictionnaryDay
        Set resultat = GetDictionnaryDay(jsonObject, month, day)
        
        ' Convertir le dictionnaire en une chaîne de caractères JSON pour l'affichage
        jsonContent = "{"
        
        ' Ajouter le commentaire du jour au contenu JSON
        If Not resultat Is Nothing Then
            jsonContent = jsonContent & """Commentary"": """ & resultat("Commentary") & """, "
            
            ' Ajouter les KeyTrade au contenu JSON
            jsonContent = jsonContent & """KeyTrade"": ["
            For Each keyTrade In resultat("KeyTrade")
                jsonContent = jsonContent & keyTrade & ", "
            Next keyTrade
            ' Supprimer la dernière virgule et l'espace supplémentaire
            If Len(jsonContent) > 1 Then
                jsonContent = Left(jsonContent, Len(jsonContent) - 2)
            End If
            jsonContent = jsonContent & "]"
        Else
            ' Si le dictionnaire est vide, ajouter des valeurs par défaut
            jsonContent = jsonContent & """Commentary"": """ & "" & """, "
            jsonContent = jsonContent & """KeyTrade"": []"
        End If
        
        ' Fermer l'objet JSON
        jsonContent = jsonContent & "}"
        
        ' Afficher le contenu JSON dans la fenêtre de débogage
        Debug.Print jsonContent
    Else
        Debug.Print "Le fichier JSON pour l'année " & year & " n'a pas été trouvé."
    End If
End Sub
