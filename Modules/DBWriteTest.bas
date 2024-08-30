Attribute VB_Name = "DBWriteTest"
' 0.5seconde sauvegarder le fichier.
' Nickel 31*0.5=15secondes Temps élevé... pour les rapports Months
Sub InsertCommentIn2024JSON()
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    
    ' Démarrer le chronomètre
    startTime = Timer
    
    ' Chemin du fichier JSON 2024
    Dim jsonFilePath As String
    jsonFilePath = ThisWorkbook.Path & "\DB\2024.json"
    
    ' Charger le contenu du fichier JSON
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(CreateObject("Scripting.FileSystemObject").OpenTextFile(jsonFilePath).ReadAll)
    
    ' Insérer le commentaire dans le mois de janvier (par exemple)
    Dim month As Object
    Set month = jsonObject("1")
    month("1")("Commentary") = "Premier commentaire de l'année 2024"
    
    ' Sauvegarder les modifications dans le fichier JSON
    SaveJSONToFile jsonObject, jsonFilePath
    
    ' Arrêter le chronomètre
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' Afficher le temps écoulé
    MsgBox "Commentaire inséré avec succès dans le fichier 2024.json." & vbCrLf & "Temps écoulé : " & elapsedTime & " secondes."
End Sub
