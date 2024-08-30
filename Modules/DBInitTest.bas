Attribute VB_Name = "DBInitTest"
Sub TestInitialization()
    ' Initialiser les fichiers JSON pour les ann�es de 2024 � 2026
    InitializeYearJSONFiles 2024, 100
    
    ' Afficher un message indiquant que les fichiers ont �t� initialis�s
    MsgBox "Les fichiers JSON ont �t� initialis�s pour les ann�es 2024 � 2026. Veuillez v�rifier le r�pertoire du classeur pour les fichiers cr��s."
End Sub

Sub TestRemoveAllJSONFiles()
    Dim fso As Object
    Dim folderPath As String
    Dim i As Integer
    Dim fileStream As Object
    
    ' Chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Cr�er l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cr�er le dossier DB s'il n'existe pas
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder (folderPath)
    End If
    
    ' Cr�er des fichiers JSON fictifs pour le test
    For i = 1 To 5
        Set fileStream = fso.CreateTextFile(folderPath & "\test" & i & ".json", True)
        fileStream.WriteLine "{""test"": ""This is a test file.""}"
        fileStream.Close
    Next i
    
    MsgBox "Fichiers JSON de test cr��s."
    
    ' Appeler la fonction pour supprimer les fichiers JSON
    RemoveAllJSONFiles
    
    ' V�rifier si les fichiers ont �t� supprim�s
    If fso.GetFolder(folderPath).Files.Count = 0 Then
        MsgBox "Tous les fichiers JSON ont �t� supprim�s avec succ�s."
    Else
        MsgBox "Certains fichiers JSON n'ont pas �t� supprim�s."
    End If
End Sub
