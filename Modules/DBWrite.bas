Attribute VB_Name = "DBWrite"
'Insertion Commentaire Month
'Insertion Commentaire year
'Insertion commentaire day
'Insertion KeyTrades


Sub InsertCommentaryDay(jsonObject As Object, month As Integer, day As Integer, Commentary As String)
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' V�rifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' Ins�rer le commentaire pour le jour sp�cifi�
            dayEntry("Commentary") = Commentary
        End If
    End If
End Sub

Sub InsertKeyTradesDay(jsonObject As Object, month As Integer, day As Integer, keyTrades As Collection)
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' V�rifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' Ins�rer les KeyTrades pour le jour sp�cifi�
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
    
    ' Cr�er l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �crire le texte JSON dans le fichier
    Set fileStream = fso.CreateTextFile(FilePath, True) ' Utilisez False pour �craser le fichier existant
    fileStream.Write jsonText
    fileStream.Close
    
   
End Sub

Sub InsertCommentaryMonth(jsonObject As Object, month As Integer, Commentary As String)
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' Ins�rer le commentaire pour le mois sp�cifi�
        monthEntry("Commentary") = Commentary
    End If
End Sub

Sub InsertCommentaryYear(jsonObject As Object, year As Integer, Commentary As String)
    ' V�rifier si l'ann�e existe
    If jsonObject.Exists(CStr(year)) Then
        Dim yearEntry As Object
        Set yearEntry = jsonObject(CStr(year))
        
        ' Ins�rer le commentaire pour l'ann�e sp�cifi�e
        For Each monthEntry In yearEntry.Items
            monthEntry("Commentary") = Commentary
        Next monthEntry
    End If
End Sub
