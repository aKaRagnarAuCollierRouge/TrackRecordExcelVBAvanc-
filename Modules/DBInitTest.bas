Attribute VB_Name = "DBInitTest"
Sub TestInitialization()
    ' Initialiser les fichiers JSON pour les années de 2024 à 2026
    InitializeYearJSONFiles 2024, 100
    
    ' Afficher un message indiquant que les fichiers ont été initialisés
    MsgBox "Les fichiers JSON ont été initialisés pour les années 2024 à 2026. Veuillez vérifier le répertoire du classeur pour les fichiers créés."
End Sub

Sub TestRemoveAllJSONFiles()
    Dim fso As Object
    Dim folderPath As String
    Dim i As Integer
    Dim fileStream As Object
    
    ' Chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Créer le dossier DB s'il n'existe pas
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder (folderPath)
    End If
    
    ' Créer des fichiers JSON fictifs pour le test
    For i = 1 To 5
        Set fileStream = fso.CreateTextFile(folderPath & "\test" & i & ".json", True)
        fileStream.WriteLine "{""test"": ""This is a test file.""}"
        fileStream.Close
    Next i
    
    MsgBox "Fichiers JSON de test créés."
    
    ' Appeler la fonction pour supprimer les fichiers JSON
    RemoveAllJSONFiles
    
    ' Vérifier si les fichiers ont été supprimés
    If fso.GetFolder(folderPath).Files.Count = 0 Then
        MsgBox "Tous les fichiers JSON ont été supprimés avec succès."
    Else
        MsgBox "Certains fichiers JSON n'ont pas été supprimés."
    End If
End Sub
