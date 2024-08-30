VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChartEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Chart As Chart
Attribute Chart.VB_VarHelpID = -1

Private Sub Chart_DblClick(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Capturer le graphique en tant qu'image
    Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' Coller l'image dans une feuille de travail temporaire
    Dim tempSheet As Worksheet
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tempSheet.Paste

    ' Enregistrer l'image
    Dim chartPic As Object
    Set chartPic = tempSheet.Shapes(1)
    chartPic.Copy

    ' Créer une image temporaire
    Dim tempChart As ChartObject
    Set tempChart = tempSheet.ChartObjects.Add(Left:=0, Width:=chartPic.Width * 1.5, Top:=0, Height:=chartPic.Height * 1.5)
    tempChart.Chart.Paste
    tempSheet.Shapes(tempChart.Name).Copy

    ' Afficher l'UserForm avec l'image agrandie
    With UserFormGraph.ImageGraph
        .Picture = Clipboard.GetData
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

    UserFormGraph.Show

    ' Supprimer la feuille de travail temporaire
    tempSheet.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
