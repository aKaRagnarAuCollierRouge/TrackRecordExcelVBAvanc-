Attribute VB_Name = "DBFetch"
Function Loadjson(year As Integer) As Object
    Dim jsonfile As String
    Dim fso As Object
    Dim fileStream As Object
    Dim jsonObject As Object
    Dim jsonString As String
    
    ' D�finir le chemin du fichier JSON de l'ann�e sp�cifi�e
    jsonfile = ThisWorkbook.Path & "\DB\" & CStr(year) & ".json"
    
    ' Cr�er l'objet FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' V�rifier si le fichier JSON existe
    If fso.FileExists(jsonfile) Then
        ' Lire le fichier JSON existant
        
        ' Ouvrir le fichier JSON en mode lecture
        Set fileStream = fso.OpenTextFile(jsonfile, 1) ' 1 = ForReading
        
        ' Lire tout le contenu du fichier JSON
        jsonString = fileStream.ReadAll
        
        ' Afficher le contenu du fichier JSON dans la fen�tre de d�bogage
        
        ' Convertir le contenu JSON en objet VBA
        Set jsonObject = JsonConverter.ParseJson(jsonString)
        
        ' Fermer le flux de fichier
        fileStream.Close
        
        ' Afficher le succ�s de chargement du fichier JSON dans la fen�tre de d�bogage
        Debug.Print "Fichier JSON charg� avec succ�s."
        
        ' Retourner l'objet JSON charg�
        Set Loadjson = jsonObject
    Else
        ' Indiquer que le fichier JSON n'existe pas
        Set Loadjson = Nothing
        Debug.Print "Fichier JSON non trouv� pour l'ann�e " & year & "."
    End If
End Function


' Dico des des Keytrades D{Commentary:...; KeyTrade:[23:....,24:....]}
'Cette fonction ne peut pas aller car elle r�cup�re tout les objets
'Moi je veux qu'elle les r�cup�re mais aussi qu'elle v�rifie que l'objet existe sinon le cr�ait avec un commentaire vide.

Function GetKeyTradesDay(jsonObject As Object, month As Integer, day As Integer) As Collection
    Dim keyTrades As New Collection
    
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' V�rifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' R�cup�rer les �l�ments de KeyTrade
            Dim keyTrade As Variant
            For Each keyTrade In dayEntry("KeyTrade")
                keyTrades.Add keyTrade
            Next keyTrade
        End If
    End If
    
    Set GetKeyTradesDay = keyTrades
End Function


Function GetCommentaryDay(ByVal jsonObject As Object, ByVal month As Integer, ByVal day As Integer) As String
    Dim Commentary As String
    
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' V�rifier si le jour existe
        If monthEntry.Exists(CStr(day)) Then
            Dim dayEntry As Object
            Set dayEntry = monthEntry(CStr(day))
            
            ' R�cup�rer le commentaire du jour
            Commentary = dayEntry("Commentary")
        End If
    End If
    
    GetCommentaryDay = Commentary
End Function

Function FindCommentMonth(ByVal jsonObject As Object, ByVal month As Integer) As String
    Dim Commentary As String
    
    ' V�rifier si le mois existe
    If jsonObject.Exists(CStr(month)) Then
        Dim monthEntry As Object
        Set monthEntry = jsonObject(CStr(month))
        
        ' R�cup�rer le commentaire du mois
        Commentary = monthEntry("Commentary") & vbCrLf
        
    End If
    
    FindCommentMonth = Commentary
End Function

Function FindCommentYear(jsonObject As Object) As String
    Dim Commentary As String
    
    Commentary = jsonObject("Commentary")
    
    FindCommentYear = Commentary
End Function

Function GetDictionnaryDayCommentary(jsonObject As Object, month As Integer, day As Integer) As Scripting.Dictionary
    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary
    
    ' V�rifier si jsonObject n'est pas Nothing
    If Not jsonObject Is Nothing Then
        ' V�rifier si le mois existe dans l'objet JSON
        If jsonObject.Exists(CStr(month)) Then
            ' V�rifier si le jour existe dans l'objet JSON pour le mois sp�cifi�
            If jsonObject(CStr(month)).Exists(CStr(day)) Then
                Set result = jsonObject(CStr(month))(CStr(day))
            End If
        End If
    End If
    
    Set GetDictionnaryDayCommentary = result
End Function

