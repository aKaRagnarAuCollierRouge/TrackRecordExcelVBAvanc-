VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Calendar"
   ClientHeight    =   9530.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14780
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private DAYBUTTON(1 To 42) As New DayButtonHandler


Public CurrentYearCalendar As Integer
Public CurrentMonthCalendar As Integer

Public dictDatasYear As Scripting.Dictionary
Public dictDatasMonth As Scripting.Dictionary
Public dictDatasDay As Scripting.Dictionary

Public dictStatsMonth As Scripting.Dictionary

Public FirstYearInitialization As Integer



Private Sub UserForm_Initialize()
    FirstYearInitialization = 2000
    Set dictDatasYear = New Scripting.Dictionary
    Set dictDatasMonth = New Scripting.Dictionary
    Set dictDatasDay = New Scripting.Dictionary
    Set dictStatsMonth = New Scripting.Dictionary
    ' Ajout des mois en anglais dans MonthList
    With Me.MonthList
        .Clear ' Supprime les �l�ments existants
        ' Ajoute les mois en anglais � la liste
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
        
        
        
    End With

    ' Ajout des ann�es dans YearList
    With Me.YearList
        .Clear ' Supprime les �l�ments existants
        ' Ajoute les ann�es � la liste
        Dim i As Integer
        For i = FirstYearInitialization To FirstYearInitialization + 200 ' Correction: Utilisation de FirstYearInitialization au lieu de i
            .AddItem i
        Next i
      
    End With

    ' Initialisation de la classe DAYBUTTON pour les boutons
  
    For i = 1 To 42
        Set DAYBUTTON(i).DAYBUTTON = Me.Controls("DayBtn" & i) ' Correction: Utilisation de Controls pour acc�der au bouton
    Next i
End Sub
'Met l'index � au mois et � la date courante.
Private Sub UserForm_Activate()
   
    
    Me.YearList.ListIndex = year(Date) - FirstYearInitialization
    Me.MonthList.ListIndex = month(Date)
    CurrentYearCalendar = year(Date)
    CurrentMonthCalendar = month(Date)
    
    
End Sub
'G�rer le changement de Mois ( passage Janvier � D�cembre)
Private Sub PreviousBtn_Click()
    ' R�cup�rer la s�lection actuelle dans les listes d�roulantes
    Dim currentYear As Integer
    Dim currentMonthIndex As Integer
    currentYear = Me.YearList.value
    currentMonthIndex = Me.MonthList.ListIndex
    Debug.Print (currentMonthIndex)
    
    ' Reculer d'un mois
    If currentMonthIndex = 0 Then
        ' Si on est en janvier, passer � d�cembre de l'ann�e pr�c�dente
        currentYear = currentYear - 1
        Me.YearList.value = currentYear
        Me.MonthList.ListIndex = 11 ' D�cembre est � l'index 11
    Else
        ' Sinon, simplement reculer d'un mois dans la m�me ann�e
        Me.MonthList.ListIndex = currentMonthIndex - 1
    End If
    
End Sub

'G�rer le changement de Mois ( passage D�cembre � Janvier)
Private Sub NextBtn_Click()
    ' R�cup�rer la s�lection actuelle dans les listes d�roulantes
    Dim currentYear As Integer
    Dim currentMonthIndex As Integer
    currentYear = Me.YearList.value
    currentMonthIndex = Me.MonthList.ListIndex
    Debug.Print (currentMonthIndex)
    ' Avancer d'un mois
    If currentMonthIndex = 11 Then
        ' Si on est en d�cembre, passer � janvier de l'ann�e suivante
        currentYear = currentYear + 1
        Me.YearList.value = currentYear
        Me.MonthList.ListIndex = 0 ' Janvier est � l'index 0
    Else
        ' Sinon, simplement avancer d'un mois dans la m�me ann�e
        Me.MonthList.ListIndex = currentMonthIndex + 1
    End If
  
End Sub

'Update jour lorsque l'on change combobox

Private Sub MonthList_Change()
    CurrentYearCalendar = Me.YearList.value
    CurrentMonthCalendar = Me.MonthList.ListIndex
    UpdateDaysAndData
    TakeCommentaryMonth Me.CurrentYearCalendar, Me.CurrentMonthCalendar + 1
    PreparationDictStatsDayAndDisplayDataMonth Calendar.CurrentYearCalendar, Calendar.CurrentMonthCalendar + 1
End Sub
'Update jour lorsque l'on change combobox


Private Sub YearList_Change()
    CurrentYearCalendar = Me.YearList.value
    CurrentMonthCalendar = Me.MonthList.ListIndex
    
    PreparationDictStatsDayAndDisplayDataMonth Calendar.CurrentYearCalendar, Calendar.CurrentMonthCalendar + 1
    TakeCommentaryMonth Me.CurrentYearCalendar, Me.CurrentMonthCalendar + 1
    UpdateDaysAndData
    
End Sub

' Fonction qui sert � changer les Days et la couleur de fond lorsque l'on change les mois et les ann�es
'NE PAS EFFACER

Private Sub UpdateDaysAndData()
    Dim year As Integer
    Dim month As Integer
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim dayBtn As MSForms.Control
    Dim trade As Variant
    Dim totalRR As Double
    Dim totalWin As Integer
    Dim totalLoose As Integer
    Dim averageRR As Double
    Dim firstDayOfMonth As Date
    Dim firstDayOfWeek As Integer
    Dim startBtnIndex As Integer
    Dim number As Integer
    
    year = YearList.value ' Assurez-vous que YearList est le nom de votre contr�le de liste des ann�es
    month = MonthList.ListIndex + 1 ' Assurez-vous que MonthList est le nom de votre contr�le de liste des mois
    
    Call PreparationDatasBtnDay(year, month)
    
    ' Calculer le premier jour du mois et le jour de la semaine correspondant
    firstDayOfMonth = DateSerial(year, month, 1)
    firstDayOfWeek = Weekday(firstDayOfMonth, vbSunday) ' vbSunday pour lundi = 1
    
    ' Calculer l'indice de d�part pour les boutons de jour
    startBtnIndex = firstDayOfWeek
    
    ' Initialiser les totaux
    totalRR = 0
    totalWin = 0
    totalLoose = 0
    
    ' Effacer le texte et r�initialiser la couleur de fond de tous les boutons de jour
    For number = 1 To 42 ' Supposons que vous avez 42 boutons de jour
        On Error Resume Next
        Set dayBtn = Me.Controls("DayBtn" & number)
        If Not dayBtn Is Nothing Then
            dayBtn.Caption = ""
            dayBtn.BackColor = RGB(255, 255, 255) ' Blanc pour neutre
        End If
        On Error GoTo 0 ' R�initialiser la gestion des erreurs � sa valeur par d�faut
    Next number
    
    ' Mettre � jour les jours et les donn�es
    For Each key In Calendar.dictDatasDay.Keys
        Set dayBtn = Me.Controls("DayBtn" & (key + startBtnIndex - 1))
        
        ' Compter les trades, calculer le RR total et mettre � jour le bouton
        Dim dayTotalRR As Double
        dayTotalRR = 0
        For Each trade In Calendar.dictDatasDay(key)("Trades")
            dayTotalRR = dayTotalRR + trade
        Next trade
        
        totalRR = totalRR + dayTotalRR
        totalWin = totalWin + Calendar.dictDatasDay(key)("nbwin")
        totalLoose = totalLoose + Calendar.dictDatasDay(key)("nbloose")
        
        ' Mettre � jour l'affichage du bouton du jour
        If dayTotalRR > 0 Then
            dayBtn.BackColor = RGB(0, 255, 0) ' Vert pour positif
        ElseIf dayTotalRR < 0 Then
            dayBtn.BackColor = RGB(255, 0, 0) ' Rouge pour n�gatif
        Else
            dayBtn.BackColor = RGB(255, 255, 255) ' Blanc pour neutre
        End If
        
        'Text+ Style Text BTN
        With dayBtn
        .Caption = key
        .Font.Size = 14 ' Changez la taille de la police selon vos besoins
        .Font.Bold = True ' Mettre le texte en gras
    End With
    Next key
    
    ' Calculer le RR moyen
    If totalWin + totalLoose > 0 Then
        averageRR = totalRR / (totalWin + totalLoose)
    Else
        averageRR = 0
    End If
    
End Sub

' DayValue & MonthValue: NbWin NbLoose WinPourcent LoosePourcent Gain Loss PourcentGain RRAverage TotalProfitRR

'Datas de Calendar.dictDatasDay

Public Sub ChangeLabelsWhenMouseMoveBtn(ByVal dayValue As String)
    '1) Extraire les donn�es de dictDatas
    Dim dictDay As Scripting.Dictionary
    
    ' Assurez-vous que dictData est d�fini et initialis� correctement
    If dictData.Exists(dayValue) Then
        Set dictDay = dictData
    Else
        Debug.Print "Aucune donn�e pour ce jour : " & dayValue
        Exit Sub
    End If
    
    ' V�rifier si le dictionnaire contient des donn�es pour ce jour
    If Not dictDay Is Nothing Then
        '2) R�cup�rer les valeurs du dictionnaire pour ce jour
        Dim NbWin As Integer
        Dim NbLoose As Integer
        
        If dictDay.Exists("nbwin") Then
            NbWin = dictDay("nbwin")
        Else
            NbWin = 0
        End If
        
        If dictDay.Exists("nbloose") Then
            NbLoose = dictDay("nbloose")
        Else
            NbLoose = 0
        End If
        
        Debug.Print "NbWin: " & NbWin
        Debug.Print "NbLoose: " & NbLoose
        
        '3) Calculer les pourcentages
        Dim WinPourcent As Double
        Dim LoosePourcent As Double
        
        If NbWin + NbLoose <> 0 Then
            WinPourcent = NbWin / (NbWin + NbLoose) * 100
            LoosePourcent = NbLoose / (NbWin + NbLoose) * 100
        Else
            WinPourcent = 0
            LoosePourcent = 0
        End If
        
        '4) Changer les labels
        If Me.Controls("NbWinDayValue") Is Nothing Then
            Debug.Print "Contr�le NbWinDayValue introuvable"
        Else
            Me.Controls("NbWinDayValue").Caption = NbWin
        End If
        
        If Me.Controls("NbLooseDayValue") Is Nothing Then
            Debug.Print "Contr�le NbLooseDayValue introuvable"
        Else
            Me.Controls("NbLooseDayValue").Caption = NbLoose
        End If
        
        If Me.Controls("PourcentWinDayValue") Is Nothing Then
            Debug.Print "Contr�le PourcentWinDayValue introuvable"
        Else
            Me.Controls("PourcentWinDayValue").Caption = Format(WinPourcent, "0.00") & "%"
        End If
        
        If Me.Controls("PourcentLooseDayValue") Is Nothing Then
            Debug.Print "Contr�le PourcentLooseDayValue introuvable"
        Else
            Me.Controls("PourcentLooseDayValue").Caption = Format(LoosePourcent, "0.00") & "%"
        End If
        
    Else
        Debug.Print "Aucune donn�e pour ce jour"
    End If
End Sub

Public Sub DisplayRapportDay()
    DayResume.Show
End Sub

Private Sub PreparationDatasYearDict(year As Integer)
    Dim jsonData As Object
    Set jsonData = Loadjson(year)
    
    ' V�rifier que jsonData est bien un dictionnaire
    If Not jsonData Is Nothing Then
        If TypeName(jsonData) = "Dictionary" Then
            Set Calendar.dictDatasYear = jsonData
            Debug.Print ("Le dictionnaire public Calendat.dictDatasYear est bien intiialis�")
        Else
            Debug.Print "Erreur: Les donn�es charg�es ne sont pas un dictionnaire valide."
        End If
    Else
        Debug.Print "Erreur: Les donn�es n'ont pas �t� charg�es pour l'ann�e " & year & "."
    End If
End Sub

Private Sub PreparationDatasMonthDict(month As Integer)
    ' V�rifier que les donn�es de l'ann�e sont charg�es
    If Not Calendar.dictDatasYear Is Nothing Then
        ' V�rifier que le mois existe dans les donn�es de l'ann�e
        If Calendar.dictDatasYear.Exists(CStr(month)) Then
            Set Calendar.dictDatasMonth = Calendar.dictDatasYear(CStr(month))
            Debug.Print ("Donn�es bien charg� dictDatasMonth")
        Else
            Debug.Print "Le mois " & month & " n'existe pas dans les donn�es de l'ann�e."
            Set Calendar.dictDatasMonth = Nothing
        End If
    Else
        Debug.Print "Les donn�es de l'ann�e ne sont pas charg�es."
        Set Calendar.dictDatasMonth = Nothing
    End If
End Sub

Public Sub PreparationDictStatsDayAndDisplayDataMonth(year As Integer, month As Integer)
    Dim DateDebut As Date
    Dim DateFin As Date
    Dim RowFiltred As Collection
    
    Dim NbWinMonth As Integer
    Dim NbLooseMonth As Integer
    Dim GainMonth As Integer
    Dim LossMonth As Integer
    Dim PourcentWinMonth As Double
    Dim PourcentLooseMonth As Double
    Dim PourcentGainMonth As Double
    Dim PourcentLossMonth As Double
    Dim RRAverageMonth As Double
    
    Dim RRMonthTotal As Double
    Dim RRGainMonth As Double
    Dim RRLooseMonth As Double
    
    Dim ColonneRR As Integer
    Dim ColonneGain As Integer
    Dim ColonneDateDebut As Integer
    Dim tableName As String
    Dim i As Integer
    Dim row As Variant
    Dim jour As Variant
    Dim RR As Double
    Dim Gain As Double
    Dim data As Scripting.Dictionary
    Dim TotalTrades As Integer
    Dim wsTrackrecord As Worksheet
    
    ' R�initialiser le dictionnaire dictStatsMonth
    Set Me.dictStatsMonth = New Scripting.Dictionary
    
    Set wsTrackrecord = ThisWorkbook.Sheets("Trackrecord")
    tableName = "Tableau1"
    ColonneRR = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "RR")
    ColonneGain = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Gain")
    ColonneDateDebut = FindIndiceCollumnWithTable(wsTrackrecord, tableName, "Date D�but")
    
    Set RowFiltred = FiltreRowByDate(DateSerial(year, month, 1), DateSerial(year, month + 1, 1), ColonneDateDebut, "Trackrecord")
    
    ' Initialiser les variables du mois
    NbWinMonth = 0
    NbLooseMonth = 0
    GainMonth = 0
    LossMonth = 0
    RRMonthTotal = 0
    RRGainMonth = 0
    RRLooseMonth = 0
    
    ' Boucle sur RowFiltred pour remplir le dictionnaire et mettre � jour les variables du mois
    For Each row In RowFiltred
        jour = day(row(1, ColonneDateDebut))
        RR = row(1, ColonneRR)
        Gain = row(1, ColonneGain)
        
        ' V�rifier si la cl� existe dans le dictionnaire
        If Not Me.dictStatsMonth.Exists(jour) Then
            Set data = New Scripting.Dictionary
            data("NbWin") = 0
            data("NbLoose") = 0
            data("Gain") = 0
            data("Loss") = 0
            data("TotalRR") = 0
            data("TotalRRGain") = 0
            data("TotalRRLoss") = 0
            Me.dictStatsMonth.Add jour, data
        Else
            Set data = Me.dictStatsMonth(jour)
        End If
        
        ' Mise � jour des donn�es journali�res
        If RR > 0 Then
            data("NbWin") = data("NbWin") + 1
            NbWinMonth = NbWinMonth + 1
            data("TotalRRGain") = data("TotalRRGain") + RR
            RRGainMonth = RRGainMonth + RR
        Else
            data("NbLoose") = data("NbLoose") + 1
            NbLooseMonth = NbLooseMonth + 1
            data("TotalRRLoss") = data("TotalRRLoss") + RR
            RRLooseMonth = RRLooseMonth + RR
        End If
        
        data("TotalRR") = data("TotalRR") + RR
        RRMonthTotal = RRMonthTotal + RR
        
        If Gain > 0 Then
            data("Gain") = data("Gain") + Gain
            GainMonth = GainMonth + Gain
        Else
            data("Loss") = data("Loss") + Gain ' Gain est n�gatif ici
            LossMonth = LossMonth + Gain ' Gain est n�gatif ici
        End If
        
        ' Mettre � jour le dictionnaire sans r�affecter l'objet entier
        Set Me.dictStatsMonth(jour) = data
    Next row
    
    ' Calcul des pourcentages pour chaque jour
    For Each jour In Me.dictStatsMonth.Keys
        Set data = Me.dictStatsMonth(jour)
        TotalTrades = data("NbWin") + data("NbLoose")
        
        If TotalTrades > 0 Then
            data("PourcentWin") = data("NbWin") / TotalTrades * 100
            data("PourcentLoose") = data("NbLoose") / TotalTrades * 100
        Else
            data("PourcentWin") = 0
            data("PourcentLoose") = 0
        End If
        
        If (data("Gain") + data("Loss")) <> 0 Then
            data("PourcentGain") = data("Gain") / (data("Gain") + data("Loss")) * 100
            data("PourcentLoss") = data("Loss") / (data("Gain") + data("Loss")) * 100
        Else
            data("PourcentGain") = 0
            data("PourcentLoss") = 0
        End If
    Next jour
    
    ' Calcul des pourcentages pour le mois
    TotalTrades = NbWinMonth + NbLooseMonth
    
    If TotalTrades > 0 Then
        PourcentWinMonth = NbWinMonth / TotalTrades * 100
        PourcentLooseMonth = NbLooseMonth / TotalTrades * 100
    Else
        PourcentWinMonth = 0
        PourcentLooseMonth = 0
    End If
    
    If (GainMonth + LossMonth) <> 0 Then
        PourcentGainMonth = GainMonth / (GainMonth + LossMonth) * 100
        PourcentLossMonth = LossMonth / (GainMonth + LossMonth) * 100
    Else
        PourcentGainMonth = 0
        PourcentLossMonth = 0
    End If
    
    ' Affichage des r�sultats mensuels
    Me.TradesTakenMonthValue = NbWinMonth + NbLooseMonth
    Me.NbWinMonthValue = NbWinMonth
    Me.NbLooseMonthValue = NbLooseMonth
    Me.GainMonthValue = GainMonth
    Me.LossMonthValue = LossMonth
    Me.PourcentWinMonthValue = PourcentWinMonth
    Me.PourcentLooseMonthValue = PourcentLooseMonth
    Me.PourcentGainMonthValue = PourcentGainMonth
    Me.PourcentLoosMonthValue = PourcentLossMonth
    Me.TotalProfitRRMonthValue = RRMonthTotal
    If NbWinMonth + NbLooseMonth > 0 Then
        Me.RRAverageMonthValue = RRMonthTotal / (NbWinMonth + NbLooseMonth)
    Else
        Me.RRAverageMonthValue = 0
    End If
    Me.RRGainMonthValue = RRGainMonth
    Me.RRLossMonthValue = RRLooseMonth
End Sub


Private Sub TakeCommentaryMonth(year As Integer, month As Integer)
    Dim fileJ As Variant
    Dim Commentary As String

    ' Load JSON data for the given year
    Set fileJ = Loadjson(year)
    
    ' Find commentary for the given month
    Commentary = FindCommentMonth(fileJ, month)
    
    ' Set the commentary month property
    Me.CommentaryMonth = Commentary
End Sub

