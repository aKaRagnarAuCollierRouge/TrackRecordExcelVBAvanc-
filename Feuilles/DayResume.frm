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

Dim imgControl As MSForms.Image ' D�clarer imgControl comme variable globale

Dim buttonHandlers As Collection
' Collection principale pour regrouper lescollections imbriqu�es
Dim controlCollections As Collection
Public DateResumeDay As Date

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl
    
    If Not ctrl Is Nothing Then
        Debug.Print "Contr�le cliqu� : " & ctrl.Name
        
        ' V�rifier si le contr�le cliqu� est un bouton
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
    ControlHeight = 200 ' Hauteur estim�e de chaque groupe de contr�les
    
    ' Initialiser la collection principale
    Set controlCollections = New Collection
    Set DatasTradesDay = PreparationResumeDay(DateResumeDay)
    Dim rowData As Scripting.Dictionary
    ' V�rifier si DatasTradesDay contient des �l�ments
    
    
    ' Appeler la fonction pour cr�er les contr�les dans la Frame
    CreateControlsWithScrolling Me.Frame, ControlHeight, DatasTradesDay
End Sub

Sub CreateControlsWithScrolling(Frame As MSForms.Frame, ControlHeight As Integer, TradesDatasCollection As Collection)
    Dim i As Long
    Dim TopPos As Long
    Dim totalHeight As Long

    ' Position de d�part
    TopPos = 10

    ' Ajuster la largeur de la Frame pour s'adapter � la largeur de la UserForm
    If Not Me Is Nothing Then
        Frame.Width = Me.Width - 30
    Else
        MsgBox "Erreur : 'Me' n'est pas d�fini."
        Exit Sub
    End If

    ' Chemin de l'image par d�faut
    Dim DefaultImagePath As String
    DefaultImagePath = ThisWorkbook.Path & "\DB\NOP.jpg"

    ' Boucle pour cr�er chaque ensemble de contr�les
    For i = 1 To TradesDatasCollection.Count
        ' R�cup�rer l'�l�ment de la collection � l'index i
        If TypeOf TradesDatasCollection(i) Is Scripting.Dictionary Then
            Dim TradeData As Scripting.Dictionary
            Set TradeData = TradesDatasCollection(i)
        Else
            MsgBox "L'�l�ment � l'index " & i & " n'est pas un Scripting.Dictionary"
            Exit Sub
        End If

        ' V�rifiez si TradeData est bien d�fini
        If TradeData Is Nothing Then
            MsgBox "Erreur : TradeData n'est pas d�fini pour l'index " & i
            Exit Sub
        End If
        
        ' Cr�er une instance de ClsResumeDayControls
        Dim tradeControls As New ClsResumeDayControls
        If tradeControls Is Nothing Then
            MsgBox "Erreur : Impossible de cr�er une instance de ClsResumeDayControls pour l'index " & i
            Exit Sub
        End If

        ' Initialise l'objet
        tradeControls.Init Frame, TradeData, TopPos, DefaultImagePath, i, ControlHeight

        ' Ajuster la position verticale pour le prochain ensemble de contr�les
        TopPos = TopPos + ControlHeight + 390 ' Ajuster en fonction des hauteurs cumul�es

        ' Calculer la hauteur totale des contr�les ajout�s
        totalHeight = totalHeight + ControlHeight + 390
    Next i
    
    ' Ajuster la hauteur de d�filement de la Frame
    Frame.ScrollHeight = totalHeight + 500 ' Ajouter une marge suppl�mentaire pour plus de visibilit�
End Sub

' Fonction pour charger l'image en fonction de la s�lection dans la combobox
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

' Fonction pour cr�er et charger l'image dans le contr�le Image en utilisant WIA
Function CreateImageControl(parentFrame As MSForms.Frame, ImagePath As String) As Object
    Dim imgControl As Object ' Utiliser un objet g�n�rique pour l'image
    Dim wiaImage As Object
    
    ' Cr�er un cadre pour l'image dans le parent sp�cifi�
    Set imgControl = parentFrame.Controls.Add("Forms.Image.1", "ImageControl")
    
    ' Si un chemin d'image est sp�cifi�, charger l'image en utilisant WIA
    If ImagePath <> "" Then
        Set wiaImage = CreateObject("WIA.ImageFile")
        wiaImage.LoadFile ImagePath
        Set imgControl.Picture = wiaImage.FileData.Picture
    End If
    
    ' D�finir les propri�t�s du contr�le Image
    With imgControl
        .Top = 0
        .Left = 0
        .Width = parentFrame.Width
        .Height = parentFrame.Height
    End With
    
    ' Retourner le contr�le Image cr�� et charg�
    Set CreateImageControl = imgControl
End Function

' Fonction pour charger une image en utilisant WIA
Function LoadImage(ByVal Filename As String) As StdPicture
    Dim wiaImage As Object
    
    ' Cr�er un objet WIA.ImageFile
    Set wiaImage = CreateObject("WIA.ImageFile")
    
    ' Charger le fichier image sp�cifi�
    wiaImage.LoadFile Filename
    
    ' R�cup�rer l'image sous forme de StdPicture
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
    tableName = "Tableau1" ' Remplacez par le nom r�el de votre tableau
    Set wsTrackrecord = ThisWorkbook.Worksheets("Trackrecord")
    Set RowFiltred = FiltreRowByDate(DateJour, DateAdd("d", 1, DateJour), 1, "Trackrecord")
    ColonneDateFin = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date Fin")
    ColonneRR = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "RR")
    ColonneGain = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Gain")
    ColonneDateDebut = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date D�but")
    ColonneHeureEntree = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure D�but")
    ColonneHeureSortie = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Heure Fin")
    ColonneKeyTrade = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "KeyTrade")
    ColonneActif = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Actif")
    Debug.Print ("Colonne RR:" & ColonneRR)
    ' Obtenir la collection des colonnes "Screenshot"
    Set screenshotCols = FindColumnsWithPattern(wsTrackrecord, tableName, "Screenshot")
    
    ' Charger le fichier JSON et obtenir le dictionnaire pour le jour sp�cifique
    Set JSONdico = LoadjsonObject(year(DateJour))
    Set dicoDay = GetDictionnaryDayCommentary(JSONdico, month(DateJour), day(DateJour))
    
    ' Initialiser les collections et les variables de comptage
    Set resultCollection = New Collection
    winCount = 0
    lossCount = 0
    totalRR = 0
    
    ' Boucler sur les lignes filtr�es
    Debug.Print ("ROwFiltred.COutn" & RowFiltred.Count)
    For i = 1 To RowFiltred.Count
        Dim rowData As Scripting.Dictionary
        Dim ligneArray As Variant
        Dim rrValue As Double
        Dim keyTrade As String
        Dim Commentary As String
        
        ligneArray = RowFiltred(i)
        
        ' Cr�er un nouveau dictionnaire pour stocker les donn�es de la ligne
        Set rowData = New Scripting.Dictionary
        rowData.Add "Date D�but", ligneArray(1, ColonneDateDebut)
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
        
        ' Obtenir le commentaire associ� au Key Trade
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
        
        ' Ajouter le dictionnaire � la collection de r�sultats
        resultCollection.Add rowData
    Next i
    
    ' Afficher les r�sultats pour v�rification
    Debug.Print "Wins: " & winCount
    Debug.Print "Losses: " & lossCount
    Debug.Print "Total RR: " & totalRR
    
    
    For Each rowData In resultCollection
        Debug.Print "Date D�but: " & rowData("Date D�but")
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




