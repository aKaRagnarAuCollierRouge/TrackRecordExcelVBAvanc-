VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthResume 
   ClientHeight    =   8030
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13670
   OleObjectBlob   =   "MonthResume.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthResume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim imgControl As MSForms.Image ' D�clarer imgControl comme variable globale

Private Sub UserForm_Initialize()
    Dim ImagePath As String
    
    ' Chemin de l'image initiale
    ImagePath = "C:\Users\Baptiste\Desktop\fichier trade en r��l\Banques des trades\2024\Juin\15\CADCHF (2).png"
    
    ' Cr�er dynamiquement un contr�le Image
    Set imgControl = Me.Controls.Add("Forms.Image.1", "DynamicImage")
    
    ' D�finir les propri�t�s du contr�le Image initial
    With imgControl
        .Top = 10 ' Position verticale
        .Left = 10 ' Position horizontale
        .Width = 200 ' Largeur souhait�e de l'image
        .Height = 50 ' Hauteur souhait�e de l'image
        .PictureSizeMode = fmPictureSizeModeStretch ' Redimensionner l'image pour remplir la zone d�finie
        .Picture = LoadImage(ImagePath) ' Charger l'image en utilisant WIA
    End With
End Sub

' Fonction pour charger une image en utilisant WIA
Function LoadImage(ByVal Filename As String) As StdPicture
    With CreateObject("WIA.ImageFile")
        .LoadFile Filename
        Set LoadImage = .FileData.Picture
    End With
End Function

' Proc�dure pour changer dynamiquement l'image
Sub ClearAndReplaceImage()
    Dim newImagePath As String
    
    ' Chemin de la nouvelle image
    newImagePath = "C:\Users\Baptiste\Desktop\fichier trade en r��l\Banques des trades\2024\Juin\15\CADCHF (3).png"
    
    ' Effacer l'image actuelle
    If Not imgControl Is Nothing Then
        imgControl.Picture = Nothing
    End If
    
    ' Charger la nouvelle image
    imgControl.Picture = LoadImage(newImagePath)
End Sub

' �v�nement associ� au clic sur le bouton
Private Sub CommandButton1_Click()
    ClearAndReplaceImage
End Sub
