Attribute VB_Name = "handles"
Function ExtraireValeurPourCle(dict As Object, cle As String) As Variant
    ' Vérifier si la clé existe dans le dictionnaire
    If dict.Exists(cle) Then
        ' Retourner la valeur associée à la clé
        ExtraireValeurPourCle = dict(cle)
    Else
        ' Retourner une valeur vide si la clé n'existe pas
        ExtraireValeurPourCle = Empty
        Debug.Print ("La Clé n'existe pas")
    End If
End Function

Function VerificationDivision0(numérateur, dénominateur)
    If dénominateur = 0 Then
        VerificationDivision0 = 0
    Else
        VerificationDivision0 = numérateur / dénominateur
    End If
    
End Function

Function IsAlpha(str As String) As Boolean
    Dim i As Integer
    IsAlpha = True
    For i = 1 To Len(str)
        If Not Mid(str, i, 1) Like "[A-Z]" Then
            IsAlpha = False
            Exit Function
        End If
    Next i
End Function

' Fonction pour vérifier si une chaîne est une lettre (vérification de colonne)
Function IsColumn(str As String) As Boolean
    Dim colNum As Long
    On Error Resume Next
    colNum = range(str & "1").Column
    IsColumn = (colNum > 0)
    On Error GoTo 0
End Function

Function FindPreviousMonthAndYear() As Scripting.Dictionary
    Dim selectedDate As Date
    Dim selectedMonth As Integer
    Dim selectedYear As Integer
    
    Dim dict As New Scripting.Dictionary
    
    ' Calcul de la date du mois dernier
    selectedDate = DateAdd("m", -1, Date)
    
    ' Extraire le mois et l'année de la date du mois dernier
    selectedMonth = month(selectedDate)
    selectedYear = year(selectedDate)
    
    dict.Add "Month", selectedMonth
    dict.Add "Year", selectedYear
    
    ' Retourner le dictionnaire
    Set FindPreviousMonthAndYear = dict
End Function

Function ColonneLettreToNuméro(ByVal lettreColonne As String) As Integer
    Dim i As Integer
    Dim résultat As Integer
    résultat = 0
    
    For i = 1 To Len(lettreColonne)
        résultat = résultat * 26 + (Asc(UCase(Mid(lettreColonne, i, 1))) - 64)
    Next i
    
    ColonneLettreToNuméro = résultat
End Function
'Okay testé
Function GetFirstAndLastDayOfMonth(ByVal year As Integer, ByVal month As Integer) As Variant
    Dim DateFirstDayMonth As Date
    Dim DateLastDayMonth As Date
    
    ' Déterminer la première journée du mois
    DateFirstDayMonth = DateSerial(year, month, 1)
    
    ' Déterminer la dernière journée du mois
    DateLastDayMonth = DateSerial(year, month + 1, 0)
    
    ' Retourner les deux dates
    GetFirstAndLastDayOfMonth = Array(DateFirstDayMonth, DateLastDayMonth)
End Function

Function CalculPourcentWhitoutBug(numerateur, denominateur)
    If denominateur = 0 Then
        CalculPourcentWithoutBug = 0
    Else
        CalculPourcentWithoutBug = numerateur / denominateur
    
    End If
End Function

Function ConvertirEnAdresseAbsolue(ByVal adress As String, ByVal sheetName As String) As String
    Dim adresseSansDollar As String
    
    ' Supprimer les symboles "$" s'ils sont présents
    adresseSansDollar = Replace(adress, "$", "")
    
    ' Ajouter le point d'exclamation et le nom de la feuille
    ConvertirEnAdresseAbsolue = sheetName & "!" & adresseSansDollar
End Function




' A mettre dans Settings Exctract Datas....'
Function FindYearLine(year As Integer) As Integer
    Dim wsSettings As Worksheet
    Dim FirstYear As Integer
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    FirstYear = wsSettings.range("A2").value
    FindYearLine = year - FirstYear + 1

End Function

Function FindMonthLine(month As Integer) As Integer
    FindMonthLine = month + 1
  
End Function

Function FindCoordonneesMonthYearRapportMonth(year As Integer, month As Integer) As Collection
   
    Dim wsSettings As Worksheet
    Dim BeginLine As Integer
    Dim EndLine As Integer
    Dim BeginCollumn As Integer
    Dim EndCollumn As Integer
    Dim YearLine As Integer
    Dim MonthLine As Integer
    Dim result As New Collection
    
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    YearLine = FindYearLine(year)
    MonthLine = FindMonthLine(month)
     
    BeginLine = wsSettings.range("B" & YearLine).value
    EndLine = wsSettings.range("C" & YearLine)
    BeginCollumn = ColonneLettreToNuméro(wsSettings.range("E" & MonthLine).value)
    EndCollumn = ColonneLettreToNuméro(wsSettings.range("F" & MonthLine).value)
 
    
    result.Add BeginLine, "BeginLine"
    result.Add EndLine, "EndLine"
    result.Add BeginCollumn, "BeginCollumn"
    result.Add EndCollumn, "EndCollumn"
    
    Set FindCoordonneesMonthYearRapportMonth = result
    
End Function

Function PointExists(series As series, pointIndex As Integer) As Boolean
    On Error Resume Next
    Dim dummy As Variant
    dummy = series.Points(pointIndex).value
    PointExists = (Err.number = 0)
    
    If PointExists Then
        Debug.Print "Point " & pointIndex & " exists."
    Else
        Debug.Print "Point " & pointIndex & " does not exist."
    End If
    
    On Error GoTo 0
End Function


