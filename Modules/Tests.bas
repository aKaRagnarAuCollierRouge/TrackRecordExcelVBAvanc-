Attribute VB_Name = "Tests"




Sub TestRangeLocation()
    Dim wsSettings As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim locationRange As range
    
    ' Message de début du test
    Debug.Print "--------------Test de la fonction RangeLocation--------------"
    
    ' Spécifiez la feuille de calcul "Settings"
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Spécifiez l'année et le mois pour le test
    year = 2025 ' Changer l'année selon vos besoins
    month = 12 ' Changer le mois selon vos besoins
    
    ' Appel de la fonction RangeLocation
    Set locationRange = RangeLocationDefine(wsSettings, year, month)
    
    ' Vérifiez si la plage a été correctement définie
    If Not locationRange Is Nothing Then
        Debug.Print "La plage de localisation a été définie avec succès."
        ' Afficher les adresses de début et de fin de la plage
        Debug.Print "Adresse de début: " & locationRange.Cells(1, 1).Address
        Debug.Print "Adresse de fin: " & locationRange.Cells(locationRange.Rows.Count, locationRange.Columns.Count).Address
    Else
        Debug.Print "La plage de localisation n'a pas été définie correctement."
    End If
End Sub

Sub TestRangeLocationDefine2()
    Dim wsParams As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim testRange As range
    
    ' Spécifiez la feuille de calcul "Settings"
    Set wsParams = ThisWorkbook.Sheets("Settings")
    
    year = 2023 ' Changer l'année selon vos besoins
    month = 5 ' Changer le mois selon vos besoins
    
    Set testRange = RangeLocationDefine(wsParams, year, month)

    ' Vérifiez si le range a été correctement défini
    If Not testRange Is Nothing Then
        Debug.Print "Le range a été défini avec succès."
        Debug.Print "La cellule supérieure gauche du range est : " & testRange.Left & " " & testRange.Top
    Else
        Debug.Print "Le range n'a pas été défini correctement."
    End If
End Sub






