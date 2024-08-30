Attribute VB_Name = "handlesTests"

Sub TestGetFirstAndLastDayOfMonth()
    Dim testCases As Variant
    Dim result As Variant
    Dim year As Integer
    Dim month As Integer
    
    ' Spécifier les cas de test : (année, mois)
    testCases = Array(Array(2024, 5), _
                      Array(2023, 12), _
                      Array(2022, 2), _
                      Array(2021, 8))
    
    ' Boucler à travers les cas de test
    For Each testCase In testCases
        year = testCase(0)
        month = testCase(1)
        
        ' Appel de la fonction GetFirstAndLastDayOfMonth
        result = GetFirstAndLastDayOfMonth(year, month)
        
        ' Afficher les résultats
        Debug.Print "Pour l'année " & year & " et le mois " & month & ":"
        Debug.Print "Premier jour du mois: " & result(0)
        Debug.Print "Dernier jour du mois: " & result(1)
        Debug.Print "----------------------"
    Next testCase
End Sub




