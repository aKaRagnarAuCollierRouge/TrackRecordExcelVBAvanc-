VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DayResume 
   Caption         =   "UserForm1"
   ClientHeight    =   10040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   17780
   OleObjectBlob   =   "DayResume.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DayResume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim imgControl As MSForms.Image ' Déclarer imgControl comme variable globale

Dim buttonHandlers As Collection
' Collection principale pour regrouper lescollections imbriquées
Dim controlCollections As Collection
Public DateResumeDay As Date

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl
    
    If Not ctrl Is Nothing Then
        Debug.Print "Contrôle cliqué : " & ctrl.Name
        
        ' Vérifier si le contrôle cliqué est un bouton
        If TypeOf ctrl Is MSForms.CommandButton Then
            Dim btn As MSForms.CommandButton
            Set btn = ctrl
            
            ' Afficher le Tag du bouton
            Debug.Print "Tag du bouton : " & btn.Tag
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    Set buttonHandlers = New Collection
    

End Sub



Private Sub UserForm_Activate()
    Dim DatasTradesDay As Collection
    Dim ControlHeight As Integer
    ControlHeight = 200 ' Hauteur estimée de chaque groupe de contrôles
    
    ' Initialiser la collection principale
    Set controlCollections = New Collection
    Set DatasTradesDay = PreparationResumeDay(DateResumeDay)
    Dim rowData As Scripting.Dictionary
    ' Vérifier si DatasTradesDay contient des éléments
    
    
    ' Appeler la fonction pour créer les contrôles dans la Frame
    CreateControlsWithScrolling Me.Frame, ControlHeight, DatasTradesDay
End Sub

Sub CreateControlsWithScrolling(Frame As MSForms.Frame, ControlHeight As Integer, TradesDatasCollection As Collection)
    Dim i As Long
    Dim TopPos As Long
    Dim totalHeight As Long

    ' Position de départ
    TopPos = 10

    ' Ajuster la largeur de la Frame pour s'adapter à la largeur de la UserForm
    If Not Me Is Nothing Then
        Frame.Width = Me.Width - 30
    Else
        MsgBox "Erreur : 'Me' n'est pas défini."
        Exit Sub
    End If

    ' Chemin de l'image par défaut
    Dim DefaultImagePath As String
    DefaultImagePath = ThisWorkbook.Path & "\DB\NOP.jpg"

    ' Boucle pour créer chaque ensemble de contrôles
    For i = 1 To TradesDatasCollection.Count
        ' Récupérer l'élément de la collection à l'index i
        If TypeOf TradesDatasCollection(i) Is Scripting.Dictionary Then
            Dim TradeData As Scripting.Dictionary
            Set TradeData = TradesDatasCollection(i)
        Else
            MsgBox "L'élément à l'index " & i & " n'est pas un Scripting.Dictionary"
            Exit Sub
        End If

        ' Vérifiez si TradeData est bien défini
        If TradeData Is Nothing Then
            MsgBox "Erreur : TradeData n'est pas défini pour l'index " & i
            Exit Sub
        End If
        
        ' Créer une instance de ClsResumeDayControls
        Dim tradeControls As New ClsResumeDayControls
        If tradeControls Is Nothing Then
            MsgBox "Erreur : Impossible de créer une instance de ClsResumeDayControls pour l'index " & i
            Exit Sub
        End If

        ' Initialise l'objet
        tradeControls.Init Frame, TradeData, TopPos, DefaultImagePath, i, ControlHeight

        ' Ajuster la position verticale pour le prochain ensemble de contrôles
        TopPos = TopPos + ControlHeight + 390 ' Ajuster en fonction des hauteurs cumulées

        ' Calculer la hauteur totale des contrôles ajoutés
        totalHeight = totalHeight + ControlHeight + 390
    Next i
    
    ' Ajuster la hauteur de défilement de la Frame
    Frame.ScrollHeight = totalHeight + 500 ' Ajouter une marge supplémentaire pour plus de visibilité
End Sub

' Fonction pour charger l'image en fonction de la sélection dans la combobox
Private Sub cboScreenshots_Change()
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.ActiveControl
    Dim imgName As String
    imgName = cbo.Tag
    Me.Controls(imgName).Picture = LoadImage(cbo.value)
End Sub

Function FileExists(FilePath As String) As Boolean
If FilePath <> "" Then
FileExists = Dir(FilePath) <> ""
Else
FileExists = False
End If
End Function

' Fonction pour créer et charger l'image dans le contrôle Image en utilisant WIA
Function CreateImageControl(parentFrame As MSForms.Frame, ImagePath As String) As Object
    Dim imgControl As Object ' Utiliser un objet générique pour l'image
    Dim wiaImage As Object
    
    ' Créer un cadre pour l'image dans le parent spécifié
    Set imgControl = parentFrame.Controls.Add("Forms.Image.1", "ImageControl")
    
    ' Si un chemin d'image est spécifié, charger l'image en utilisant WIA
    If ImagePath <> "" Then
        Set wiaImage = CreateObject("WIA.ImageFile")
        wiaImage.LoadFile ImagePath
        Set imgControl.Picture = wiaImage.FileData.Picture
    End If
    
    ' Définir les propriétés du contrôle Image
    With imgControl
        .Top = 0
        .Left = 0
        .Width = parentFrame.Width
        .Height = parentFrame.Height
    End With
    
    ' Retourner le contrôle Image créé et chargé
    Set CreateImageControl = imgControl
End Function

' Fonction pour charger une image en utilisant WIA
Function LoadImage(ByVal Filename As String) As StdPicture
    Dim wiaImage As Object
    
    ' Créer un objet WIA.ImageFile
    Set wiaImage = CreateObject("WIA.ImageFile")
    
    ' Charger le fichier image spécifié
    wiaImage.LoadFile Filename
    
    ' Récupérer l'image sous forme de StdPicture
    Set LoadImage = wiaImage.FileData.Picture
End Function

Function PreparationResumeDay(DateJour As Date) As Collection
    Dim RowFiltred As Collection
    Dim wsTrackrecord As Worksheet
    Dim tableName As String
    Dim ColonneRR As Integer
    Dim ColonneDateDebut As Integer
    Dim ColonneDateFin As Integer
    Dim ColonneHeureEntree As Integer
    Dim ColonneHeureSortie As Integer
    Dim ColonneKeyTrade As Integer
    Dim ColonneGain As Integer
    Dim ColonneActif As Integer
    Dim JSONdico As Object
    Dim dicoDay As Scripting.Dictionary
    Dim resultCollection As Collection
    Dim screenshotCols As Collection
    Dim i As Integer
    Dim winCount As Integer
    Dim lossCount As Integer
    Dim totalRR As Double
    Debug.Print ("DatePrepa:" & DateJour)
    tableName = "Tableau1" ' Remplacez par le nom réel de votre tableau
    Set wsTrackrecord = ThisWorkbook.Worksheets("Trackrecord")
    Set RowFiltred = FiltreRowByDate(DateJour, DateAdd("d", 1, DateJour), 1, "Trackrecord")
    ColonneDateFin = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date Fin")
    ColonneRR = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "RR")
    ColonneGain = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Gain")
    ColonneDateDebut = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date Début")
    ColonneHeureEntree = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure Début")
    ColonneHeureSortie = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure Fin")
    ColonneKeyTrade = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "KeyTrade")
    ColonneActif = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Actif")
    Debug.Print ("Colonne RR:" & ColonneRR)
    ' Obtenir la collection des colonnes "Screenshot"
    Set screenshotCols = FindColumnsWithPattern(wsTrackrecord, tableName, "Screenshot")
    
    ' Charger le fichier JSON et obtenir le dictionnaire pour le jour spécifique
    Set JSONdico = LoadjsonObject(year(DateJour))
    Set dicoDay = GetDictionnaryDayCommentary(JSONdico, month(DateJour), day(DateJour))
    
    ' Initialiser les collections et les variables de comptage
    Set resultCollection = New Collection
    winCount = 0
    lossCount = 0
    totalRR = 0
    
    ' Boucler sur les lignes filtrées
    Debug.Print ("ROwFiltred.COutn" & RowFiltred.Count)
    For i = 1 To RowFiltred.Count
        Dim rowData As Scripting.Dictionary
        Dim ligneArray As Variant
        Dim rrValue As Double
        Dim keyTrade As String
        Dim Commentary As String
        
        ligneArray = RowFiltred(i)
        
        ' Créer un nouveau dictionnaire pour stocker les données de la ligne
        Set rowData = New Scripting.Dictionary
        rowData.Add "Date Début", ligneArray(1, ColonneDateDebut)
        rowData.Add "Date Fin", ligneArray(1, ColonneDateFin)
        rowData.Add "Heure Entree", ligneArray(1, ColonneHeureEntree)
        rowData.Add "Heure Sortie", ligneArray(1, ColonneHeureSortie)
        rowData.Add "RR", ligneArray(1, ColonneRR)
        rowData.Add "Gain", ligneArray(1, ColonneGain)
        rowData.Add "KeyTrade", ligneArray(1, ColonneKeyTrade)
        rowData.Add "Actif", ligneArray(1, ColonneActif)
        
        ' Ajouter les valeurs des colonnes "Screenshot" au dictionnaire
        Dim colIndex As Variant
        For Each colIndex In screenshotCols
            rowData.Add "Screenshot" & colIndex, ligneArray(1, colIndex)
        Next colIndex
        
        rrValue = ligneArray(1, ColonneRR)
        keyTrade = ligneArray(1, ColonneKeyTrade)
        Debug.Print ("RRvalue:" & rrValue)
        Debug.Print ("keytrade:" & keyTrade)
        
        ' Obtenir le commentaire associé au Key Trade
        If dicoDay.Exists(keyTrade) Then
            Commentary = dicoDay(keyTrade)("Commentary")
        Else
            Commentary = ""
        End If
        rowData.Add "Commentary", Commentary
        
        ' Compter les wins et les losses
        If rrValue > 0 Then
            winCount = winCount + 1
        ElseIf rrValue < 0 Then
            lossCount = lossCount + 1
        End If
        
        ' Ajouter le RR au total
        totalRR = totalRR + rrValue
        
        ' Ajouter le dictionnaire à la collection de résultats
        resultCollection.Add rowData
    Next i
    
    ' Afficher les résultats pour vérification
    Debug.Print "Wins: " & winCount
    Debug.Print "Losses: " & lossCount
    Debug.Print "Total RR: " & totalRR
    
    
    For Each rowData In resultCollection
        Debug.Print "Date Début: " & rowData("Date Début")
        Debug.Print "Date Fin: " & rowData("Date Fin")
        Debug.Print "Heure Entree: " & rowData("Heure Entree")
        Debug.Print "Heure Sortie: " & rowData("Heure Sortie")
        Debug.Print "RR: " & rowData("RR")
        Debug.Print "Gain: " & rowData("Gain")
        Debug.Print "Key Trade: " & rowData("KeyTrade")
        Debug.Print "Commentary: " & rowData("Commentary")
        Debug.Print "Actif: " & rowData("Actif")
        For Each colIndex In screenshotCols
            Debug.Print "Screenshot" & colIndex & ": " & rowData("Screenshot" & colIndex)
        Next colIndex
        Debug.Print "-------------------------------"
    Next rowData
    
    Set PreparationResumeDay = resultCollection
End Function




