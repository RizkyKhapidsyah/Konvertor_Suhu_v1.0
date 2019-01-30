VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FormNolAbsolut 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nol Absolut"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormNolAbsolut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl MesinXP 
      Left            =   240
      Top             =   2520
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin MSComctlLib.ListView LV 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FormNolAbsolut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Kriteria", 2100
        .ColumnHeaders.Add , , "Kelvin", 1000
        .ColumnHeaders.Add , , "Celcius", 1000
        .ColumnHeaders.Add , , "Fahrenheit", 1000
        .ColumnHeaders.Add , , "Rankine", 1000
        .ColumnHeaders.Add , , "Delisle", 1000
        .ColumnHeaders.Add , , "Newton", 1000
        .ColumnHeaders.Add , , "Reamur", 1000
        .ColumnHeaders.Add , , "Romer", 1000
        .View = lvwReport
        .GridLines = True
    End With
    LV.ListItems.Clear
    Set LI = LV.ListItems.Add(, , "Nol Absolut")
        LI.SubItems(1) = "0 K"
        LI.SubItems(2) = "-273,15 °C"
        LI.SubItems(3) = "-459,67 °F"
        LI.SubItems(4) = "0 °Ra"
        LI.SubItems(5) = "559.73 °De"
        LI.SubItems(6) = "-90,14 °N"
        LI.SubItems(7) = "-218,52 °Ré"
        LI.SubItems(8) = "-135,9 °Rø"
    Set LI = LV.ListItems.Add(, , "Titik Beku Air")
        LI.SubItems(1) = "273,15 K"
        LI.SubItems(2) = "0 °C"
        LI.SubItems(3) = "32 °F"
        LI.SubItems(4) = "491,67 °Ra"
        LI.SubItems(5) = "150 °De"
        LI.SubItems(6) = "0 °N"
        LI.SubItems(7) = "0 °Ré"
        LI.SubItems(8) = "7,5 °Rø"
    Set LI = LV.ListItems.Add(, , "Suhu Badan Manusia")
        LI.SubItems(1) = "310,15 K"
        LI.SubItems(2) = "37 °C"
        LI.SubItems(3) = "98,6 °F"
        LI.SubItems(4) = "558,27 °Ra"
        LI.SubItems(5) = "94,5 °De"
        LI.SubItems(6) = "12,21 °N"
        LI.SubItems(7) = "29,6 °Ré"
        LI.SubItems(8) = "26,93 °Rø"
    Set LI = LV.ListItems.Add(, , "Titik Didih Air")
        LI.SubItems(1) = "373,15 K"
        LI.SubItems(2) = "100 °C"
        LI.SubItems(3) = "212 °F"
        LI.SubItems(4) = "671,67 °Ra"
        LI.SubItems(5) = "0 °De"
        LI.SubItems(6) = "33 °N"
        LI.SubItems(7) = "80 °Ré"
        LI.SubItems(8) = "60 °Rø"
        MesinXP.StartEngine
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

