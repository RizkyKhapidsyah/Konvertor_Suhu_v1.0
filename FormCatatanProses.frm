VERSION 5.00
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCatatanProses 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catatan Hasil"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormCatatanProses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin XPEngine.XPControl MesinXP 
      Left            =   360
      Top             =   5760
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7435
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
Attribute VB_Name = "FormCatatanProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub MasukkanDataKeLV()
    With LV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Nilai_Awal", 2000
        .ColumnHeaders.Add , , "Hasil_1", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_2", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_3", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_4", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_5", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_6", 1200, vbCenter
        .ColumnHeaders.Add , , "Hasil_7", 1200, vbCenter
        .View = lvwReport
        .GridLines = True
        .Sorted = True
    End With
        LV.ListItems.Clear
        Do Until FormUtama.AdodcCatatanHasil.Recordset.EOF
        Set LI = LV.ListItems.Add(, , FormUtama.AdodcCatatanHasil.Recordset.Fields(0).Value)
            LI.SubItems(1) = FormUtama.AdodcCatatanHasil.Recordset.Fields(1).Value
            LI.SubItems(2) = FormUtama.AdodcCatatanHasil.Recordset.Fields(2).Value
            LI.SubItems(3) = FormUtama.AdodcCatatanHasil.Recordset.Fields(3).Value
            LI.SubItems(4) = FormUtama.AdodcCatatanHasil.Recordset.Fields(4).Value
            LI.SubItems(5) = FormUtama.AdodcCatatanHasil.Recordset.Fields(5).Value
            LI.SubItems(6) = FormUtama.AdodcCatatanHasil.Recordset.Fields(6).Value
            LI.SubItems(7) = FormUtama.AdodcCatatanHasil.Recordset.Fields(7).Value
            FormUtama.AdodcCatatanHasil.Recordset.MoveNext
        Loop
        FormUtama.AdodcCatatanHasil.Refresh
End Sub

Private Sub Form_Load()
    MasukkanDataKeLV
    MesinXP.StartEngine
End Sub

