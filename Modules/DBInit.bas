Attribute VB_Name = "DBInit"
Sub InitializeYearJSONFiles(startYear As Integer, numYears As Integer)
    Dim fso As Object
    Dim jsonObject As Object
    Dim monthEntry As Object
    Dim yearEntry As Object
    Dim dayEntry As Object
    Dim jsonText As String
    Dim year As Integer, month As Integer, day As Integer
    Dim yearFile As String
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Boucler sur chaque année pour créer un fichier JSON distinct
    For year = startYear To startYear + numYears - 1
        ' Créer l'objet JSON pour l'année
        Set jsonObject = CreateObject("Scripting.Dictionary")
        
        
        jsonObject.Add "Commentary", ""
        
        ' Ajouter les mois à l'objet JSON
        For month = 1 To 12
            Set monthEntry = CreateObject("Scripting.Dictionary")
            monthEntry("Commentary") = ""
            
            ' Ajouter les jours à chaque mois
            For day = 1 To 31 ' Vous pouvez ajuster la limite supérieure selon vos besoins
                Set dayEntry = CreateObject("Scripting.Dictionary")
                dayEntry("Commentary") = ""
                
                ' Initialiser KeyTrade comme une collection vide
                Dim keyTrade As Collection
                Set keyTrade = New Collection
                Set dayEntry("KeyTrade") = keyTrade
                
                ' Ajouter le jour au mois
                monthEntry.Add CStr(day), dayEntry
            Next day
            
            ' Ajouter le mois à l'année
            jsonObject.Add month, monthEntry
        Next month
        
        Dim i As Integer
        Dim QDic As Scripting.Dictionary
        For i = 1 To 4
            Set QDic = CreateObject("Scripting.Dictionary")
            QDic("Commentary") = ""
            jsonObject.Add "Q" & CStr(i), QDic
        Next i
        ' Convertir l'objet JSON en texte JSON
        jsonText = JsonConverter.ConvertToJson(jsonObject)
        
        ' Définir le nom du fichier JSON pour cette année
        yearFile = ThisWorkbook.Path & "\DB\" & CStr(year) & ".json"
        
        ' Écrire le texte JSON dans le fichier
        Dim fileStream As Object
        Set fileStream = fso.CreateTextFile(yearFile, True) ' Utilisez True pour écraser le fichier existant
        fileStream.Write jsonText
        fileStream.Close
    Next year
    
    MsgBox "Les fichiers JSON ont été initialisés pour les années " & startYear & " à " & startYear + numYears - 1 & "."
End Sub




Sub RemoveAllJSONFiles()
    Dim fso As Object
    Dim folderPath As String
    Dim file As Object
    
    ' Chemin du dossier contenant les fichiers JSON
    folderPath = ThisWorkbook.Path & "\DB"
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Vérifier si le dossier existe
    If fso.FolderExists(folderPath) Then
        ' Boucler sur tous les fichiers du dossier
        For Each file In fso.GetFolder(folderPath).Files
            ' Vérifier si le fichier est un fichier JSON
            If LCase(fso.GetExtensionName(file.Name)) = "json" Then
                ' Supprimer le fichier
                fso.DeleteFile file.Path
            End If
        Next file
        MsgBox "Tous les fichiers JSON ont été supprimés du dossier DB."
    Else
        MsgBox "Le dossier DB n'existe pas."
    End If
End Sub

Sub ReinitialiserDB(ByVal startYear As Integer)
    RemoveAllJSONFiles
    InitializeYearJSONFiles startYear, 100
End Sub
