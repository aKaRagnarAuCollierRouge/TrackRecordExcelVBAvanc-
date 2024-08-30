Attribute VB_Name = "DBFetchTest"
Sub TestLoadJSONObject()
    Dim yearExisting As Integer
    Dim yearNonExisting As Integer
    Dim jsonObject As Object
    Dim fso As Object
    Dim jsonfile As String
    
    ' Cr�er l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ann�e avec un fichier JSON existant
    yearExisting = 2024
    jsonfile = ThisWorkbook.Path & "\DB\" & CStr(yearExisting) & ".json"
    Debug.Print (jsonfile)
    ' V�rifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        Debug.Print "Fichier JSON existant trouv� pour l'ann�e " & yearExisting
        Set jsonObject = Loadjson(yearExisting)
        If Not jsonObject Is Nothing Then
            Debug.Print "Test avec fichier JSON existant (ann�e " & yearExisting & "):"
            ' Afficher le contenu JSON dans la fen�tre de d�bogage
            PrintJSONObject jsonObject
        Else
            Debug.Print "Test avec fichier JSON existant (ann�e " & yearExisting & "): Chargement JSON �chou�."
        End If
    Else
        Debug.Print "Fichier JSON non trouv� pour l'ann�e " & yearExisting
    End If
    
    ' Ann�e avec un fichier JSON non existant
    yearNonExisting = 2025
    jsonfile = ThisWorkbook.Path & "\DB\" & CStr(yearNonExisting) & ".json"
    
    ' V�rifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        Debug.Print "Fichier JSON existant trouv� pour l'ann�e " & yearNonExisting
        Set jsonObject = Loadjson(yearNonExisting)
        If Not jsonObject Is Nothing Then
            Debug.Print "Test avec fichier JSON non existant (ann�e " & yearNonExisting & "):"
            ' Afficher le contenu JSON dans la fen�tre de d�bogage
            PrintJSONObject jsonObject
        Else
            Debug.Print "Test avec fichier JSON non existant (ann�e " & yearNonExisting & "): Chargement JSON �chou�."
        End If
    Else
        Debug.Print "Fichier JSON non trouv� pour l'ann�e " & yearNonExisting
    End If
End Sub

' Fonction pour afficher le contenu JSON dans la fen�tre de d�bogage
Sub PrintJSONObject(jsonObject As Object)
    Dim key As Variant
    
    For Each key In jsonObject.Keys
        ' Afficher chaque cl� et sa valeur dans la fen�tre de d�bogage
        Debug.Print key & ": " & jsonObject(key)
    Next key
End Sub



Sub TestFolderExists()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print (GlobalsVariables.folderPathCommentary)
    
    ' V�rifier si le chemin existe
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
    
    ' Chemin du fichier � v�rifier
    FilePath = GlobalsVariables.folderPathCommentary + "\2024.json"
    Debug.Print (FilePath)
    ' V�rifier si le fichier existe
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
    
    ' D�finir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' D�finir l'ann�e, le mois et le jour de test
    year = 2024
    month = 1
    day = 1
    
    ' Charger le fichier JSON de l'ann�e sp�cifi�e
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et r�cup�rer les KeyTrades
    Set keyTrades = GetKeyTradesDay(jsonObject, month, day)
    
    ' Afficher les KeyTrades dans une bo�te de message
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
    
    ' D�finir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' D�finir l'ann�e, le mois et le jour de test
    year = 2024
    month = 1
    day = 1
    
    ' Charger le fichier JSON de l'ann�e sp�cifi�e
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et r�cup�rer le commentaire du jour
    Commentary = GetCommentaryDay(jsonObject, month, day)
    
    ' Afficher le commentaire dans une bo�te de message
    MsgBox "Commentaire : " & vbCrLf & Commentary, vbInformation, "Test GetCommentaryDay"
End Sub

Sub TestFindCommentMonth()
    Dim folderPath As String
    Dim year As Integer
    Dim month As Integer
    Dim result As String
    Dim jsonObject As Object
    Dim jsonfile As String
    
    ' D�finir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' D�finir l'ann�e et le mois de test
    year = 2024
    month = 1
    
    ' Charger le fichier JSON de l'ann�e sp�cifi�e
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et r�cup�rer le r�sultat
    result = FindCommentMonth(jsonObject, month)
    
    ' Afficher le r�sultat dans une bo�te de message
    MsgBox result, vbInformation, "Test FindCommentByYearMonth"
End Sub

Sub TestFindCommentYear()
    Dim folderPath As String
    Dim year As Integer
    Dim result As String
    Dim jsonObject As Object
    Dim jsonfile As String
    
    ' D�finir le chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' D�finir l'ann�e de test
    year = 2024
    
    ' Charger le fichier JSON de l'ann�e sp�cifi�e
    jsonfile = folderPath & "\" & year & ".json"
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonfile).ReadAll)
    
    ' Appeler la fonction et r�cup�rer le r�sultat
    result = FindCommentYear(jsonObject)
    
    ' Afficher le r�sultat dans une bo�te de message
    MsgBox result, vbInformation, "Test FindCommentByYear"
End Sub



Option Explicit

' D�claration de la variable globale pour le chemin du dossier
Public folderPath As String

' Cette proc�dure sera ex�cut�e � l'ouverture du classeur
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

' Fonction pour charger un objet JSON � partir d'un fichier pour une ann�e sp�cifique
Function LoadjsonObject(year As Integer) As Object
    Dim jsonfile As String
    Dim fso As Object
    Dim fileStream As Object
    Dim jsonObject As Object
    Dim FileExists As Variant

    ' D�finir le chemin du fichier JSON de l'ann�e sp�cifi�e
    jsonfile = folderPath & "\" & year & ".json"
    
    ' Cr�er l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' V�rifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        ' Lire le fichier JSON existant
        Set fileStream = fso.OpenTextFile(jsonfile, 1) ' 1 = ForReading
        Set jsonObject = JsonConverter.ParseJson(fileStream.ReadAll)
        fileStream.Close
        
        ' Indiquer que le fichier JSON a �t� charg� avec succ�s
        Set LoadjsonObject = jsonObject
    Else
        ' Indiquer que le fichier JSON n'existe pas
        Set LoadjsonObject = Nothing
    End If
End Function

' Fonction pour r�cup�rer les donn�es pour un mois et un jour sp�cifiques � partir d'un objet JSON
Function GetDictionnaryDay(jsonObject As Object, month As Integer, day As Integer) As Dictionary
    Dim result As Dictionary
    Set result = New Dictionary
    
    ' V�rifier si le mois existe dans l'objet JSON
    If jsonObject.Exists(CStr(month)) Then
        ' V�rifier si le jour existe dans l'objet JSON pour le mois sp�cifi�
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
    
    ' Sp�cifiez l'ann�e, le mois et le jour que vous voulez tester
    year = 2024
    month = 1
    day = 1
    
    ' Charger l'objet JSON pour l'ann�e sp�cifi�e
    Set jsonObject = LoadjsonObject(year)
    
    ' V�rifier si l'objet JSON a �t� charg�
    If Not jsonObject Is Nothing Then
        ' Appeler la fonction GetDictionnaryDay
        Set resultat = GetDictionnaryDay(jsonObject, month, day)
        
        ' Convertir le dictionnaire en une cha�ne de caract�res JSON pour l'affichage
        jsonContent = "{"
        
        ' Ajouter le commentaire du jour au contenu JSON
        If Not resultat Is Nothing Then
            jsonContent = jsonContent & """Commentary"": """ & resultat("Commentary") & """, "
            
            ' Ajouter les KeyTrade au contenu JSON
            jsonContent = jsonContent & """KeyTrade"": ["
            For Each keyTrade In resultat("KeyTrade")
                jsonContent = jsonContent & keyTrade & ", "
            Next keyTrade
            ' Supprimer la derni�re virgule et l'espace suppl�mentaire
            If Len(jsonContent) > 1 Then
                jsonContent = Left(jsonContent, Len(jsonContent) - 2)
            End If
            jsonContent = jsonContent & "]"
        Else
            ' Si le dictionnaire est vide, ajouter des valeurs par d�faut
            jsonContent = jsonContent & """Commentary"": """ & "" & """, "
            jsonContent = jsonContent & """KeyTrade"": []"
        End If
        
        ' Fermer l'objet JSON
        jsonContent = jsonContent & "}"
        
        ' Afficher le contenu JSON dans la fen�tre de d�bogage
        Debug.Print jsonContent
    Else
        Debug.Print "Le fichier JSON pour l'ann�e " & year & " n'a pas �t� trouv�."
    End If
End Sub
