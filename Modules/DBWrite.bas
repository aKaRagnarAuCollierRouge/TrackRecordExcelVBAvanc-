Attribute VB_Name = "DBWrite"
'Insertion Commentaire Month
'Insertion Commentaire year
'Insertion commentaire day
'Insertion KeyTrades


Sub InsertCommentaryDay(jsonObject As Object, month As Integer, day As Integer, Commentary As String)
    ' Vérifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' Vérifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' Insérer le commentaire pour le jour spécifié
            dayEntry("Commentary") = Commentary
        End If
    End If
End Sub

Sub InsertKeyTradesDay(jsonObject As Object, month As Integer, day As Integer, keyTrades As Collection)
    ' Vérifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' Vérifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' Insérer les KeyTrades pour le jour spécifié
            Dim keyTrade As Variant
            For Each keyTrade In keyTrades
                dayEntry("KeyTrade").Add keyTrade
            Next keyTrade
        End If
    End If
End Sub


Sub SaveJSONToFile(jsonObject As Object, FilePath As String)
    Dim fso As Object
    Dim fileStream As Object
    Dim jsonText As String
    
    ' Convertir l'objet JSON en texte JSON
    jsonText = JsonConverter.ConvertToJson(jsonObject)
    
    ' Créer l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Écrire le texte JSON dans le fichier
    Set fileStream = fso.CreateTextFile(FilePath, True) ' Utilisez False pour écraser le fichier existant
    fileStream.Write jsonText
    fileStream.Close
    
   
End Sub

Sub InsertCommentaryMonth(jsonObject As Object, month As Integer, Commentary As String)
    ' Vérifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' Insérer le commentaire pour le mois spécifié
        monthEntry("Commentary") = Commentary
    End If
End Sub

Sub InsertCommentaryYear(jsonObject As Object, year As Integer, Commentary As String)
    ' Vérifier si l'année existe
    If jsonObject.Exists(CStr(year)) Then
        Dim yearEntry As Object
        Set yearEntry = jsonObject(CStr(year))
        
        ' Insérer le commentaire pour l'année spécifiée
        For Each monthEntry In yearEntry.Items
            monthEntry("Commentary") = Commentary
        Next monthEntry
    End If
End Sub
