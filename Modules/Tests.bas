Attribute VB_Name = "Tests"




Sub TestRangeLocation()
    Dim wsSettings As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim locationRange As range
    
    ' Message de d�but du test
    Debug.Print "--------------Test de la fonction RangeLocation--------------"
    
    ' Sp�cifiez la feuille de calcul "Settings"
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    ' Sp�cifiez l'ann�e et le mois pour le test
    year = 2025 ' Changer l'ann�e selon vos besoins
    month = 12 ' Changer le mois selon vos besoins
    
    ' Appel de la fonction RangeLocation
    Set locationRange = RangeLocationDefine(wsSettings, year, month)
    
    ' V�rifiez si la plage a �t� correctement d�finie
    If Not locationRange Is Nothing Then
        Debug.Print "La plage de localisation a �t� d�finie avec succ�s."
        ' Afficher les adresses de d�but et de fin de la plage
        Debug.Print "Adresse de d�but: " & locationRange.Cells(1, 1).Address
        Debug.Print "Adresse de fin: " & locationRange.Cells(locationRange.Rows.Count, locationRange.Columns.Count).Address
    Else
        Debug.Print "La plage de localisation n'a pas �t� d�finie correctement."
    End If
End Sub

Sub TestRangeLocationDefine2()
    Dim wsParams As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim testRange As range
    
    ' Sp�cifiez la feuille de calcul "Settings"
    Set wsParams = ThisWorkbook.Sheets("Settings")
    
    year = 2023 ' Changer l'ann�e selon vos besoins
    month = 5 ' Changer le mois selon vos besoins
    
    Set testRange = RangeLocationDefine(wsParams, year, month)

    ' V�rifiez si le range a �t� correctement d�fini
    If Not testRange Is Nothing Then
        Debug.Print "Le range a �t� d�fini avec succ�s."
        Debug.Print "La cellule sup�rieure gauche du range est : " & testRange.Left & " " & testRange.Top
    Else
        Debug.Print "Le range n'a pas �t� d�fini correctement."
    End If
End Sub






