Attribute VB_Name = "DBRemove"
Function RemoveKeyTradesDay(jsonObject As Object, month As Integer, day As Integer, keyTradesToRemove As Collection)
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' V�rifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' R�cup�rer les �l�ments de KeyTrade existants
            Dim existingKeyTrades As Collection
            Set existingKeyTrades = New Collection
            
            Dim keyTrade As Variant
            For Each keyTrade In dayEntry("KeyTrade")
                existingKeyTrades.Add keyTrade
            Next keyTrade
            
            ' Supprimer les KeyTrade sp�cifi�s
            Dim i As Long
            For i = existingKeyTrades.Count To 1 Step -1
                Dim keyTradeEntry As String
                keyTradeEntry = existingKeyTrades(i)
                
                ' V�rifier si le KeyTrade est dans la liste des KeyTrade � supprimer
                Dim keyTradeToRemove As Variant
                For Each keyTradeToRemove In keyTradesToRemove
                    If InStr(keyTradeEntry, keyTradeToRemove & ":") = 1 Then
                        existingKeyTrades.Remove i
                        Exit For
                    End If
                Next keyTradeToRemove
            Next i
            
            ' Mettre � jour les KeyTrade dans l'objet JSON
            Set dayEntry("KeyTrade") = existingKeyTrades
        End If
    End If
End Function
