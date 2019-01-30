VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FormRingkasan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ringkasan"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormRingkasan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textSkala 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FormRingkasan.frx":000C
      Top             =   720
      Width           =   3975
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   7440
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.ComboBox cmbSkala 
      Height          =   390
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin isButton3.isButton cmTutup 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Icon            =   "FormRingkasan.frx":0012
      Style           =   7
      Caption         =   "&Tutup"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skala     :"
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FormRingkasan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
Attribute AturKontrol.VB_UserMemId = -550
    With Me
        .cmbSkala.Clear
        .cmbSkala.AddItem "Kelvin", 0
        .cmbSkala.AddItem "Celcius", 1
        .cmbSkala.AddItem "Fahrenheit", 2
        .cmbSkala.AddItem "Rankine", 3
        .cmbSkala.AddItem "Delisle", 4
        .cmbSkala.AddItem "Newton", 5
        .cmbSkala.AddItem "Réamur", 6
        .cmbSkala.AddItem "Rømer", 7
        .cmbSkala.AddItem "Leiden", 8
        .cmbSkala.ListIndex = 0
        .textSkala.Text = ""
        .textSkala.Locked = True
    End With
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Konversi ke : ", 1140
        .ColumnHeaders.Add , , "Rumus ", 2840
        .View = lvwReport
        .GridLines = True
    End With
End Sub

Public Sub cmbSkala_Click()
On Error GoTo PecahkanError
    LV.ListItems.Clear
    Select Case cmbSkala.ListIndex
        Case Is = 0
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "K - 273,15"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "K × 1,8 - 459,67"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "K × 1,8"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(373,15 - K) × 1,5"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "(K - 273,15) × 33/100"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "(K - 273,15) × 0,8"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "(K - 273,15) × 21/40 + 7,5"
        Case Is = 1
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "°C + 273,15"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "°C × 1,8 + 32"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "°C × 1,8 + 491,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(100 - °C) × 1,5"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "°C × 33/100"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "°C × 0,8"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "°C × 21/40 + 7,5"
        Case Is = 2
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "(°F + 459,67) / 1,4"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "(°F - 32) / 1,8"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "°F + 459,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(212 - °F) × 5/6"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "(°F - 32) × 11/60"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "(°F - 32) / 2,25"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "(°F - 32) × 7/24 + 7,5"
        Case Is = 3
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "°Ra / 1,8"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "°Ra / 1,8 + 273,15"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "°Ra - 459,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(671,67 - °Ra) × 5/6"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "(°Ra - 491,67) × 11/60"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "(°Ra / 1,8 + 273,15) × 0,8"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "(°Ra - 491,67) × 7/24 + 7,5"
        Case Is = 4
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "373,15 - °De × 2/3"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "100 - °De × 2/3"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "212 - °De × 1,2"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "671,67 - °De × 1,2"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "33 - °De × 0,22"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "80 - °De × 8/15"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "60 - °De × 0,35"
        Case Is = 5
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "°N × 100/33 + 273,15"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "°N × 100/33"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "°N x 60/11 + 32"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "°N × 60/11 + 491,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(33 - °N) × 50/11"
            Set LI = LV.ListItems.Add(, , "Reamur")
                LI.SubItems(1) = "°N × 80/33"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "°N × 35/22 + 7,5"
        Case Is = 6
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "°Ré / 0,8 + 273,15"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "°Ré / 0,8"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "°Ré × 2,25 + 32"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "°Ré × 2,25 + 491,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(80 - °Ré) × 1,875"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "°Ré × 33/80"
            Set LI = LV.ListItems.Add(, , "Romer")
                LI.SubItems(1) = "°Ré × 21/32 + 7,5"
        Case Is = 7
            Set LI = LV.ListItems.Add(, , "Kelvin")
                LI.SubItems(1) = "(°Rø - 7,5) × 40/21 + 273.15"
            Set LI = LV.ListItems.Add(, , "Celcius")
                LI.SubItems(1) = "(°Rø - 7,5) × 40/21"
            Set LI = LV.ListItems.Add(, , "Fahrenheit")
                LI.SubItems(1) = "(°Rø - 7,5) × 24/7 + 32"
            Set LI = LV.ListItems.Add(, , "Rankine")
                LI.SubItems(1) = "(°Rø - 7,5) × 24/7 + 491,67"
            Set LI = LV.ListItems.Add(, , "Delisle")
                LI.SubItems(1) = "(60 - °Rø) × 20/7"
            Set LI = LV.ListItems.Add(, , "Newton")
                LI.SubItems(1) = "(°Rø - 7,5) × 22/35"
            Set LI = LV.ListItems.Add(, , "Reaumur")
                LI.SubItems(1) = "(°Rø - 7,5) × 32/21"
        Case Is = 8
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
            Set LI = LV.ListItems.Add(, , "Tidak Ada Deskripsi")
                LI.SubItems(1) = "Tidak Ada Deskripsi"
    End Select
    
    Dim i As Integer
    Dim s As String, s1 As String
    i = FreeFile

    textSkala.Text = ""
    If cmbSkala.ListIndex = 0 Then
        Open App.Path & "\referensi\rksKelvin.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 1 Then
        Open App.Path & "\referensi\rksCelcius.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 2 Then
        Open App.Path & "\referensi\rksFahrenheit.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 3 Then
        Open App.Path & "\referensi\rksRankine.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 4 Then
        Open App.Path & "\referensi\rksDelisle.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 5 Then
        Open App.Path & "\referensi\rksNewton.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 6 Then
        Open App.Path & "\referensi\rksReamur.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 7 Then
        Open App.Path & "\referensi\rksRomer.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    ElseIf cmbSkala.ListIndex = 8 Then
        Open App.Path & "\referensi\rksLeiden.rcf" For Input As #i
            Do Until EOF(i)
                Input #i, s
                s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
            Loop
        Close #i
        textSkala.Text = s1
    End If

Exit Sub
PecahkanError:
    PusatError
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
    MesinXP.StartEngine
End Sub

