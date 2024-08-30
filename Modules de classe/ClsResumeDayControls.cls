VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsResumeDayControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public WithEvents mFrame As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
Public mTradeData As Scripting.Dictionary
Public mControlSet As Collection
Public mDefaultImagePath As String
Public mTopPos As Integer
Public mIndex As Long
Public mControlHeight As Integer

Public mButtons As Collection  ' Collection pour stocker les boutons créés
Public Sub Init(Frame As MSForms.Frame, TradeData As Scripting.Dictionary, ByVal TopPos As Integer, DefaultImagePath As String, Index As Long, ControlHeight As Integer)
    ' Définir les membres de la classe
    Set Me.mFrame = Frame
    Set Me.mTradeData = TradeData
    Me.mTopPos = TopPos
    Me.mDefaultImagePath = DefaultImagePath
    Me.mIndex = Index
    Me.mControlHeight = ControlHeight
    
    ' Initialiser les collections
    Set Me.mControlSet = New Collection
    Set Me.mButtons = New Collection
    
    ' Créer les contrôles
    CreateControls
End Sub

Private Sub CreateControls()
    ' Créer les labels
    Dim lblDateDebut As MSForms.Label
    Set lblDateDebut = CreateLabel("LabelDateDebut_" & mIndex, "Date Début: " & Format(mTradeData("Date Début"), "dd/mm/yyyy"), 10, mTopPos, 130, 20)
    mControlSet.Add lblDateDebut

    Dim lblDateFin As MSForms.Label
    Set lblDateFin = CreateLabel("LabelDateFin_" & mIndex, "Date Fin: " & Format(mTradeData("Date Fin"), "dd/mm/yyyy"), 150, mTopPos, 130, 20)
    mControlSet.Add lblDateFin

    Dim lblHeureEntree As MSForms.Label
    Set lblHeureEntree = CreateLabel("LabelHeureEntree_" & mIndex, "Heure Entrée: " & Format(mTradeData("Heure Entrée"), "hh:mm:ss"), 290, mTopPos, 100, 20)
    mControlSet.Add lblHeureEntree

    Dim lblHeureSortie As MSForms.Label
    Set lblHeureSortie = CreateLabel("LabelHeureSortie_" & mIndex, "Heure Sortie: " & Format(mTradeData("Heure Sortie"), "hh:mm:ss"), 400, mTopPos, 100, 20)
    mControlSet.Add lblHeureSortie

    Dim lblActif As MSForms.Label
    Set lblActif = CreateLabel("LabelActif_" & mIndex, "Actif: " & mTradeData("Actif"), 510, mTopPos, 100, 20)
    mControlSet.Add lblActif

    ' Créer les boutons de navigation
    Dim btnPrev As MSForms.CommandButton
    Set btnPrev = CreateButton("ButtonPrev_" & mIndex, "<<", 10, mTopPos + 30, 50, 20, 1)
    mControlSet.Add btnPrev

    ' Créer la combobox pour les screenshots
    Dim cboScreenshots As MSForms.ComboBox
    Set cboScreenshots = CreateComboBox("ComboBoxScreenshots_" & mIndex, 70, mTopPos + 30, 120, 30)
    PopulateComboBoxWithScreenshots cboScreenshots
    mControlSet.Add cboScreenshots

    ' Créer le bouton suivant
    Dim btnNext As MSForms.CommandButton
    Set btnNext = CreateButton("ButtonNext_" & mIndex, ">>", 200, mTopPos + 30, 50, 20, 2)
    mControlSet.Add btnNext

    ' Créer le cadre pour l'image
    Dim imgControl As MSForms.Image
    Set imgControl = CreateImage("ImageControl_" & mIndex, 10, mTopPos + 70, 800, 300)
    imgControl.Picture = LoadImage(mDefaultImagePath)
    mControlSet.Add imgControl
End Sub

Private Function CreateLabel(Name As String, Caption As String, Left As Integer, Top As Integer, Width As Integer, Height As Integer) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = mFrame.Controls.Add("Forms.Label.1", Name)
    With lbl
        .Caption = Caption
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .Font.Name = "Arial"
        .Font.Size = 12
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(0, 102, 204)
    End With
    Set CreateLabel = lbl
End Function

Private Function CreateButton(Name As String, Caption As String, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Index As Integer) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = mFrame.Controls.Add("Forms.CommandButton.1", Name)
    With btn
        .Caption = Caption
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .Font.Name = "Arial"
        .Font.Size = 12
        .Tag = Index  ' Utilisation de la propriété Tag pour stocker l'index
    End With
    ' Ajouter le bouton à la collection de boutons
    mButtons.Add btn
    Set CreateButton = btn
End Function

Private Function CreateComboBox(Name As String, Left As Integer, Top As Integer, Width As Integer, Height As Integer) As MSForms.ComboBox
    Dim cbo As MSForms.ComboBox
    Set cbo = mFrame.Controls.Add("Forms.ComboBox.1", Name)
    With cbo
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        .Font.Name = "Arial"
        .Font.Size = 12
    End With
    Set CreateComboBox = cbo
End Function

Private Function CreateImage(Name As String, Left As Integer, Top As Integer, Width As Integer, Height As Integer) As MSForms.Image
    Dim img As MSForms.Image
    Set img = mFrame.Controls.Add("Forms.Image.1", Name)
    With img
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With
    Set CreateImage = img
End Function

Private Sub PopulateComboBoxWithScreenshots(cboScreenshots As MSForms.ComboBox)
    ' Ajouter les options de screenshots à la combobox
    cboScreenshots.AddItem mDefaultImagePath ' Ajouter le chemin de l'image par défaut
    Dim screenshotKey As Variant
    For Each screenshotKey In mTradeData.Keys
        If Left(screenshotKey, 10) = "Screenshot" Then
            cboScreenshots.AddItem mTradeData(screenshotKey)
        End If
    Next screenshotKey
End Sub

Private Function LoadImage(ByVal Filename As String) As StdPicture
    Dim wiaImage As Object
    Set wiaImage = CreateObject("WIA.ImageFile")
    wiaImage.LoadFile Filename
    Set LoadImage = wiaImage.FileData.Picture
End Function

Public Sub UpdateImage()
    Dim screenshotPath As String
    screenshotPath = cboScreenshots.List(cboScreenshots.ListIndex)

    If FileExists(screenshotPath) Then
        imgControl.Picture = LoadImage(screenshotPath)
    Else
        imgControl.Picture = LoadImage(mDefaultImagePath)
    End If
End Sub

Private Function FileExists(FilePath As String) As Boolean
    FileExists = (Dir(FilePath) <> "")
End Function

Private Sub CommandButton_Click()
    ' Vérifier si le contrôle cliqué est un bouton
    If TypeOf ctrl Is MSForms.CommandButton Then
        Dim btn As MSForms.CommandButton
        Set btn = ctrl
        
        ' Récupérer l'index du bouton à partir de la propriété Tag
        Dim btnIndex As Integer
        btnIndex = btn.Tag
        
        ' Afficher l'index dans la fenêtre Immediate (Ctrl + G pour afficher la fenêtre Immediate)
        Debug.Print "Index du bouton cliqué : " & btnIndex
        
        ' Autres actions à effectuer en fonction de l'index récupéré
        Select Case btnIndex
            Case 1
                Debug.Print "Bouton Précédent cliqué"
                ' Traitement pour le bouton précédent
            Case 2
                Debug.Print "Bouton Suivant cliqué"
                ' Traitement pour le bouton suivant
            ' Ajouter d'autres cas selon les besoins
        End Select
    End If
End Sub
