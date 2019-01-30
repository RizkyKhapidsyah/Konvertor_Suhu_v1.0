VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Begin VB.Form FormUtama 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RikySoft Konvertor Suhu 1.0"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormUtama.frx":000C
   ScaleHeight     =   6630
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdodcCatatanHasil 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin isButton3.isButton cmProses 
      Height          =   495
      Left            =   5040
      TabIndex        =   33
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormUtama.frx":213DE
      Style           =   7
      Caption         =   "&Proses"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hasil"
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4815
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   6
         Left            =   1680
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   5
         Left            =   1680
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   4
         Left            =   1680
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   3
         Left            =   1680
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   2
         Left            =   1680
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox textHasil 
         Height          =   390
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   6
         Left            =   4320
         TabIndex        =   32
         Top             =   3240
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   5
         Left            =   4320
         TabIndex        =   31
         Top             =   2760
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   4
         Left            =   4320
         TabIndex        =   30
         Top             =   2280
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   3
         Left            =   4320
         TabIndex        =   29
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   2
         Left            =   4320
         TabIndex        =   28
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   1
         Left            =   4320
         TabIndex        =   27
         Top             =   840
         Width           =   180
      End
      Begin VB.Label LabelSatuan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "---"
         Height          =   270
         Index           =   0
         Left            =   4320
         TabIndex        =   26
         Top             =   360
         Width           =   180
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   600
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   13
         Top             =   3240
         Width           =   45
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   12
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   11
         Top             =   2280
         Width           =   45
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   45
      End
      Begin VB.Label LabelHasil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "----------"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   240
      Top             =   7200
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nilai Awal"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6255
      Begin VB.ComboBox cmbSatuanNilaiAwal 
         Height          =   390
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox textNilaiAwal 
         Height          =   390
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label LabelNilaiAwal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dari                       :"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1350
      End
   End
   Begin isButton3.isButton cmReset 
      Height          =   495
      Left            =   5040
      TabIndex        =   34
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormUtama.frx":213FA
      Style           =   7
      Caption         =   "&Reset"
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
   Begin isButton3.isButton cmKeluar 
      Height          =   495
      Left            =   5040
      TabIndex        =   35
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormUtama.frx":21416
      Style           =   7
      Caption         =   "&Keluar"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   16711680
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
   Begin isButton3.isButton cmRumus 
      Height          =   495
      Left            =   5040
      TabIndex        =   36
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormUtama.frx":21432
      Style           =   7
      Caption         =   "&Rumus"
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
   Begin isButton3.isButton cmExport 
      Height          =   495
      Left            =   5040
      TabIndex        =   37
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormUtama.frx":2144E
      Style           =   7
      Caption         =   "&Export"
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
   Begin VB.Menu menuData 
      Caption         =   "Data"
      Begin VB.Menu menuRingkasan 
         Caption         =   "Ringkasan"
         Begin VB.Menu menuKelvin 
            Caption         =   "Kelvin"
         End
         Begin VB.Menu menuCelcius 
            Caption         =   "Celcius"
         End
         Begin VB.Menu menuFahrenheit 
            Caption         =   "Fahrenheit"
         End
         Begin VB.Menu menuRankine 
            Caption         =   "Rankine"
         End
         Begin VB.Menu menuDelisle 
            Caption         =   "Delisle"
         End
         Begin VB.Menu menuNewton 
            Caption         =   "Newton"
         End
         Begin VB.Menu menuReamur 
            Caption         =   "Réaumur"
         End
         Begin VB.Menu menuRomer 
            Caption         =   "Rømer"
         End
         Begin VB.Menu split1 
            Caption         =   "-"
         End
         Begin VB.Menu menuNA 
            Caption         =   "Nol Absolut"
         End
      End
      Begin VB.Menu formCatatanHasil 
         Caption         =   "Catatan Hasil"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu menuAlat 
      Caption         =   "Alat"
      Begin VB.Menu MenuPengaturan 
         Caption         =   "Pengaturan"
      End
   End
   Begin VB.Menu MenuBantuan 
      Caption         =   "Bantuan"
      Begin VB.Menu menuTentang 
         Caption         =   "Tentang"
      End
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    NyambungUtama
    With AdodcCatatanHasil
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from CatatanHasil"
        .Refresh
    End With
    MesinXP.StartEngine
        With cmbSatuanNilaiAwal
            .Clear
            .AddItem "Celcius (°C)", 0
            .AddItem "Fahrenheit (°F)", 1
            .AddItem "Réamur (°R)", 2
            .AddItem "Kelvin (K)", 3
            .AddItem "Rankine (°Ra)", 4
            .AddItem "Delisle (°De)", 5
            .AddItem "Newton (°N)", 6
            .AddItem "Rømer (°Rø)", 7
            .ListIndex = 0
        End With
    Reset
End Sub
Sub Reset()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = vbWhite
                .Locked = False
                .Text = ""
            End With
        End If
    Next
    With textHasil
        .Item(0).Locked = True
        .Item(1).Locked = True
        .Item(2).Locked = True
        .Item(3).Locked = True
        .Item(4).Locked = True
        .Item(5).Locked = True
        .Item(6).Locked = True
    End With
End Sub
Sub KunciInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Locked = True
            End With
        End If
    Next
End Sub
Sub BukaInput()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Locked = False
            End With
        End If
    Next
End Sub
Sub SimpanDataHasilCatatanKeDatabase()
    With AdodcCatatanHasil
        .Recordset.AddNew
        .Recordset.Fields(0).Value = textNilaiAwal.Text & " " & cmbSatuanNilaiAwal.Text
        .Recordset.Fields(1).Value = textHasil(0).Text & " " & LabelSatuan(0).Caption
        .Recordset.Fields(2).Value = textHasil(1).Text & " " & LabelSatuan(1).Caption
        .Recordset.Fields(3).Value = textHasil(2).Text & " " & LabelSatuan(2).Caption
        .Recordset.Fields(4).Value = textHasil(3).Text & " " & LabelSatuan(3).Caption
        .Recordset.Fields(5).Value = textHasil(4).Text & " " & LabelSatuan(4).Caption
        .Recordset.Fields(6).Value = textHasil(5).Text & " " & LabelSatuan(5).Caption
        .Recordset.Fields(7).Value = textHasil(6).Text & " " & LabelSatuan(6).Caption
        .Recordset.Update
        .Refresh
        .Refresh
    End With
End Sub


Private Sub cmbSatuanNilaiAwal_Click()
     Select Case cmbSatuanNilaiAwal.ListIndex
        Case Is = 0
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Fahrenheit"
                .Item(2).Caption = "Rankine"
                .Item(3).Caption = "Delisle"
                .Item(4).Caption = "Newton"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°F"
                    .Item(2).Caption = "°Ra"
                    .Item(3).Caption = "°De"
                    .Item(4).Caption = "°N"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 1
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Rankine"
                .Item(3).Caption = "Delisle"
                .Item(4).Caption = "Newton"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°Ra"
                    .Item(3).Caption = "°De"
                    .Item(4).Caption = "°N"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 2
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Fahrenheit"
                .Item(3).Caption = "Rankine"
                .Item(4).Caption = "Delisle"
                .Item(5).Caption = "Newton"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°F"
                    .Item(3).Caption = "°Ra"
                    .Item(4).Caption = "°De"
                    .Item(5).Caption = "°N"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 3
            With LabelHasil
                .Item(0).Caption = "Celcius"
                .Item(1).Caption = "Fahrenheit"
                .Item(2).Caption = "Rankine"
                .Item(3).Caption = "Delisle"
                .Item(4).Caption = "Newton"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°C"
                    .Item(1).Caption = "°F"
                    .Item(2).Caption = "°Ra"
                    .Item(3).Caption = "°De"
                    .Item(4).Caption = "°N"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 4
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Fahrenheit"
                .Item(3).Caption = "Delisle"
                .Item(4).Caption = "Newton"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°F"
                    .Item(3).Caption = "°De"
                    .Item(4).Caption = "°N"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 5
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Fahrenheit"
                .Item(3).Caption = "Rankine"
                .Item(4).Caption = "Newton"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°F"
                    .Item(3).Caption = "°Ra"
                    .Item(4).Caption = "°N"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 6
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Fahrenheit"
                .Item(3).Caption = "Rankine"
                .Item(4).Caption = "Delisle"
                .Item(5).Caption = "Réamur"
                .Item(6).Caption = "Rømer"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°F"
                    .Item(3).Caption = "°Ra"
                    .Item(4).Caption = "°De"
                    .Item(5).Caption = "°Ré"
                    .Item(6).Caption = "°Rø"
                End With
        Case Is = 7
            With LabelHasil
                .Item(0).Caption = "Kelvin"
                .Item(1).Caption = "Celcius"
                .Item(2).Caption = "Fahrenheit"
                .Item(3).Caption = "Rankine"
                .Item(4).Caption = "Delisle"
                .Item(5).Caption = "Newton"
                .Item(6).Caption = "Réamur"
            End With
                With LabelSatuan
                    .Item(0).Caption = "°K"
                    .Item(1).Caption = "°C"
                    .Item(2).Caption = "°F"
                    .Item(3).Caption = "°Ra"
                    .Item(4).Caption = "°De"
                    .Item(5).Caption = "°N"
                    .Item(6).Caption = "°Ré"
                End With
    End Select
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            If Objek.BackColor = Me.BackColor Then
                cmRumus_Click
            Else
                Objek.Text = ""
            End If
        End If
    Next
End Sub

Private Sub cmExport_Click()
    On Error GoTo ErrorHandler
    If textNilaiAwal.Text = "" Then
        MsgBox "Tidak bisa export, Harap isi nilai awal", vbExclamation + vbOKOnly, ""
        textNilaiAwal.SetFocus
    Else
            CommonDialog1.DialogTitle = "Export Data"
            CommonDialog1.FileName = ""
            CommonDialog1.Filter = "RikySoft Catatan File (*.rcf)|*.rcf|Microsoft Word 2003 Document (*.doc)|*doc*|Microsoft Excel 2003 Document (*.xls)|*.xls|Microsoft Rich Text Format Document (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"
            DefaultFormat
            CommonDialog1.ShowSave
            CommonDialog1.FileName = CommonDialog1.FileName
        Dim iFile As Integer
        Dim SaveFileFromTB As Boolean
        Dim TxtBox As Object
        Dim FilePath As String
        Dim Append As Boolean
        iFile = FreeFile
            If Append Then
                Open CommonDialog1.FileName For Append As #iFile
            Else
                Open CommonDialog1.FileName For Output As #iFile
            End If
        Print #iFile, "================================================================================"
        Print #iFile, LabelNilaiAwal.Caption & " " & textNilaiAwal.Text & " " & cmbSatuanNilaiAwal.Text
        Print #iFile, "================================================================================"
        Print #iFile, LabelHasil.Item(0).Caption & "    : " & textHasil.Item(0).Text & " " & LabelSatuan.Item(0).Caption
        Print #iFile, LabelHasil.Item(1).Caption & "    : " & textHasil.Item(1).Text & " " & LabelSatuan.Item(1).Caption
        Print #iFile, LabelHasil.Item(2).Caption & "    : " & textHasil.Item(2).Text & " " & LabelSatuan.Item(2).Caption
        Print #iFile, LabelHasil.Item(3).Caption & "    : " & textHasil.Item(3).Text & " " & LabelSatuan.Item(3).Caption
        Print #iFile, LabelHasil.Item(4).Caption & "    : " & textHasil.Item(4).Text & " " & LabelSatuan.Item(4).Caption
        Print #iFile, LabelHasil.Item(5).Caption & "    : " & textHasil.Item(5).Text & " " & LabelSatuan.Item(5).Caption
        Print #iFile, LabelHasil.Item(6).Caption & "    : " & textHasil.Item(6).Text & " " & LabelSatuan.Item(6).Caption
        Print #iFile, "================================================================================"
        
        SaveFileFromTB = True
    End If
ErrorHandler:
    Close #iFile
End Sub

Private Sub cmKeluar_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Pesan = MsgBox("Apakah Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If Pesan = vbYes Then
        End
    Else
        Cancel = 1
    End If
End Sub


Private Sub MenuKeluar_Click()
    Pesan = MsgBox("Apakah Anda yakin ingin keluar?", vbQuestion + vbYesNo, "Keluar?")
    If Pesan = vbYes Then End
End Sub

Public Sub cmProses_Click()
    If textNilaiAwal.Text = "" Then
        MsgBox "Silahkan isi nilai yang akan di konversi!", vbExclamation + vbOKOnly, ""
    Else
        If FormPengaturan.CekBulankanHasilPecahan.Value = Unchecked Then
            Select Case cmbSatuanNilaiAwal.ListIndex
                Case Is = 0
                    With textHasil
                        .Item(0).Text = Val(textNilaiAwal.Text) + 273.15
                        .Item(1).Text = Val(textNilaiAwal.Text) * 1.8 + 32
                        .Item(2).Text = Val(textNilaiAwal.Text) * 1.8 + 491.67
                        .Item(3).Text = (100 - Val(textNilaiAwal.Text)) * 1.5
                        .Item(4).Text = Val(textNilaiAwal.Text) * 33 / 100
                        .Item(5).Text = Val(textNilaiAwal.Text) * 0.8
                        .Item(6).Text = Val(textNilaiAwal.Text) * 21 / 40 + 7.5
                    End With
                Case Is = 1
                    With textHasil
                        .Item(0).Text = (Val(textNilaiAwal.Text) + 459.67) / 1.4
                        .Item(1).Text = (Val(textNilaiAwal.Text) - 32) / 1.8
                        .Item(2).Text = Val(textNilaiAwal.Text) + 459.67
                        .Item(3).Text = (212 - Val(textNilaiAwal.Text)) * 5 / 6
                        .Item(4).Text = (Val(textNilaiAwal.Text) - 32) * 11 / 60
                        .Item(5).Text = (Val(textNilaiAwal.Text) - 32) / 2.25
                        .Item(6).Text = (Val(textNilaiAwal.Text) - 32) * 7 / 24 + 7.5
                    End With
                Case Is = 2
                    With textHasil
                        .Item(0).Text = Val(textNilaiAwal.Text) / 0.8 + 273.15
                        .Item(1).Text = Val(textNilaiAwal.Text) / 0.8
                        .Item(2).Text = Val(textNilaiAwal.Text) * 2.25 + 32
                        .Item(3).Text = Val(textNilaiAwal.Text) * 2.25 + 491.67
                        .Item(4).Text = (80 - Val(textNilaiAwal.Text)) * 1.875
                        .Item(5).Text = Val(textNilaiAwal.Text) * 33 / 80
                        .Item(6).Text = Val(textNilaiAwal.Text) * 21 / 32 + 7.5
                    End With
                Case Is = 3
                    With textHasil
                        .Item(0).Text = Val(textNilaiAwal.Text) - 273.15
                        .Item(1).Text = Val(textNilaiAwal.Text) * 1.8 - 459.67
                        .Item(2).Text = Val(textNilaiAwal.Text) * 1.8
                        .Item(3).Text = (373.15 - Val(textNilaiAwal.Text)) * 1.5
                        .Item(4).Text = (Val(textNilaiAwal.Text) - 273.15) * 33 / 100
                        .Item(5).Text = (Val(textNilaiAwal.Text) - 273.15) * 0.8
                        .Item(6).Text = (Val(textNilaiAwal.Text) - 273.15) * 21 / 40 + 7.5
                    End With
                Case Is = 4
                    With textHasil
                        .Item(0).Text = Val(textNilaiAwal.Text) / 1.8
                        .Item(1).Text = Val(textNilaiAwal.Text) / 1.8 + 273.15
                        .Item(2).Text = Val(textNilaiAwal.Text) - 459.67
                        .Item(3).Text = (671.67 - Val(textNilaiAwal.Text)) * 5 / 6
                        .Item(4).Text = (Val(textNilaiAwal.Text) - 491.67) * 11 / 60
                        .Item(5).Text = (Val(textNilaiAwal.Text) / 1.8 + 273.15) * 0.8
                        .Item(6).Text = (Val(textNilaiAwal.Text) - 491.67) * 7 / 24 + 7.5
                    End With
                Case Is = 5
                    With textHasil
                        .Item(0).Text = 373.15 - Val(textNilaiAwal.Text) * 2 / 3
                        .Item(1).Text = 100 - Val(textNilaiAwal.Text) * 2 / 3
                        .Item(2).Text = 212 - Val(textNilaiAwal.Text) * 1.2
                        .Item(3).Text = 671.67 - Val(textNilaiAwal.Text) * 1.2
                        .Item(4).Text = 33 - Val(textNilaiAwal.Text) * 0.22
                        .Item(5).Text = 80 - Val(textNilaiAwal.Text) * 8 / 15
                        .Item(6).Text = 60 - Val(textNilaiAwal.Text) * 0.35
                    End With
                Case Is = 6
                    With textHasil
                        .Item(0).Text = Val(textNilaiAwal.Text) * 100 / 33 + 273.15
                        .Item(1).Text = Val(textNilaiAwal.Text) * 100 / 33
                        .Item(2).Text = Val(textNilaiAwal.Text) * 60 / 11 + 32
                        .Item(3).Text = Val(textNilaiAwal.Text) * 60 / 11 + 491.67
                        .Item(4).Text = (33 - Val(textNilaiAwal.Text)) * 50 / 11
                        .Item(5).Text = Val(textNilaiAwal.Text) * 80 / 33
                        .Item(6).Text = Val(textNilaiAwal.Text) * 35 / 22 + 7.5
                    End With
                Case Is = 7
                    With textHasil
                        .Item(0).Text = (Val(textNilaiAwal.Text) - 7.5) * 40 / 21 + 273.15
                        .Item(1).Text = (Val(textNilaiAwal.Text) - 7.5) * 40 / 21
                        .Item(2).Text = (Val(textNilaiAwal.Text) - 7.5) * 24 / 7 + 32
                        .Item(3).Text = (Val(textNilaiAwal.Text) - 7.5) * 24 / 7 + 491.67
                        .Item(4).Text = (60 - Val(textNilaiAwal.Text)) * 20 / 7
                        .Item(5).Text = (Val(textNilaiAwal.Text) - 7.5) * 22 / 35
                        .Item(6).Text = (Val(textNilaiAwal.Text) - 7.5) * 32 / 21
                    End With
            End Select
        ElseIf FormPengaturan.CekBulankanHasilPecahan.Value = Checked Then
            Select Case cmbSatuanNilaiAwal.ListIndex
                Case Is = 0
                    With textHasil
                        .Item(0).Text = Round(Val(textNilaiAwal.Text) + 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(Val(textNilaiAwal.Text) * 1.8 + 32, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) * 1.8 + 491.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round((100 - Val(textNilaiAwal.Text)) * 1.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round(Val(textNilaiAwal.Text) * 33 / 100, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round(Val(textNilaiAwal.Text) * 0.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round(Val(textNilaiAwal.Text) * 21 / 40 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 1
                    With textHasil
                        .Item(0).Text = Round((Val(textNilaiAwal.Text) + 459.67) / 1.4, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round((Val(textNilaiAwal.Text) - 32) / 1.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) + 459.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round((212 - Val(textNilaiAwal.Text)) * 5 / 6, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((Val(textNilaiAwal.Text) - 32) * 11 / 60, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round((Val(textNilaiAwal.Text) - 32) / 2.25, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round((Val(textNilaiAwal.Text) - 32) * 7 / 24 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 2
                    With textHasil
                        .Item(0).Text = Round(Val(textNilaiAwal.Text) / 0.8 + 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(Val(textNilaiAwal.Text) / 0.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) * 2.25 + 32, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round(Val(textNilaiAwal.Text) * 2.25 + 491.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((80 - Val(textNilaiAwal.Text)) * 1.875, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round(Val(textNilaiAwal.Text) * 33 / 80, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round(Val(textNilaiAwal.Text) * 21 / 32 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 3
                    With textHasil
                        .Item(0).Text = Round(Val(textNilaiAwal.Text) - 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(Val(textNilaiAwal.Text) * 1.8 - 459.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) * 1.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round((373.15 - Val(textNilaiAwal.Text)) * 1.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((Val(textNilaiAwal.Text) - 273.15) * 33 / 100, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round((Val(textNilaiAwal.Text) - 273.15) * 0.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round((Val(textNilaiAwal.Text) - 273.15) * 21 / 40 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 4
                    With textHasil
                        .Item(0).Text = Round(Val(textNilaiAwal.Text) / 1.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(Val(textNilaiAwal.Text) / 1.8 + 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) - 459.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round((671.67 - Val(textNilaiAwal.Text)) * 5 / 6, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((Val(textNilaiAwal.Text) - 491.67) * 11 / 60, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round((Val(textNilaiAwal.Text) / 1.8 + 273.15) * 0.8, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round((Val(textNilaiAwal.Text) - 491.67) * 7 / 24 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 5
                    With textHasil
                        .Item(0).Text = Round(373.15 - Val(textNilaiAwal.Text) * 2 / 3, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(100 - Val(textNilaiAwal.Text) * 2 / 3, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(212 - Val(textNilaiAwal.Text) * 1.2, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round(671.67 - Val(textNilaiAwal.Text) * 1.2, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round(33 - Val(textNilaiAwal.Text) * 0.22, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round(80 - Val(textNilaiAwal.Text) * 8 / 15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round(60 - Val(textNilaiAwal.Text) * 0.35, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 6
                    With textHasil
                        .Item(0).Text = Round(Val(textNilaiAwal.Text) * 100 / 33 + 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round(Val(textNilaiAwal.Text) * 100 / 33, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round(Val(textNilaiAwal.Text) * 60 / 11 + 32, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round(Val(textNilaiAwal.Text) * 60 / 11 + 491.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((33 - Val(textNilaiAwal.Text)) * 50 / 11, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round(Val(textNilaiAwal.Text) * 80 / 33, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = Round(Val(textNilaiAwal.Text) * 35 / 22 + 7.5, Val(FormPengaturan.cmbDibelakangKoma.Text))
                    End With
                Case Is = 7
                    With textHasil
                        .Item(0).Text = Round((Val(textNilaiAwal.Text) - 7.5) * 40 / 21 + 273.15, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(1).Text = Round((Val(textNilaiAwal.Text) - 7.5) * 40 / 21, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(2).Text = Round((Val(textNilaiAwal.Text) - 7.5) * 24 / 7 + 32, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(3).Text = Round((Val(textNilaiAwal.Text) - 7.5) * 24 / 7 + 491.67, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(4).Text = Round((60 - Val(textNilaiAwal.Text)) * 20 / 7, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(5).Text = Round((Val(textNilaiAwal.Text) - 7.5) * 22 / 35, Val(FormPengaturan.cmbDibelakangKoma.Text))
                        .Item(6).Text = (Val(textNilaiAwal.Text) - 7.5) * 32 / 21
                    End With
            End Select
        End If
            
        SimpanDataHasilCatatanKeDatabase
    End If
End Sub

Private Sub cmReset_Click()
    Reset
    cmProses.Enabled = True
    textNilaiAwal.SetFocus
End Sub

Public Sub cmRumus_Click()
    With textNilaiAwal
        .Text = cmbSatuanNilaiAwal.Text
        .BackColor = Me.BackColor
    End With
    Select Case cmbSatuanNilaiAwal.ListIndex
        Case Is = 0
            With textHasil
                .Item(0).Text = "°C + 273,15"
                .Item(1).Text = "°C × 1,8 + 32"
                .Item(2).Text = "°C × 1,8 + 491,67"
                .Item(3).Text = " (100 - °C) × 1,5"
                .Item(4).Text = "°C × 33/100"
                .Item(5).Text = "°C × 0,8"
                .Item(6).Text = "°C × 21/40 + 7,5"
            End With
        Case Is = 1
            With textHasil
                .Item(0).Text = "(°F + 459,67) / 1,4"
                .Item(1).Text = "(°F - 32) / 1,8"
                .Item(2).Text = "°F + 459,67"
                .Item(3).Text = "(212 - °F) × 5/6"
                .Item(4).Text = "(°F - 32) × 11/60"
                .Item(5).Text = "(°F - 32) / 2,25"
                .Item(6).Text = "(°F - 32) × 7/24 + 7,5"
            End With
        Case Is = 2
            With textHasil
                .Item(0).Text = "°Ré / 0,8 + 273,15"
                .Item(1).Text = "°Ré / 0,8"
                .Item(2).Text = "°Ré × 2,25 + 32"
                .Item(3).Text = "°Ré × 2,25 + 491,67"
                .Item(4).Text = "(80 - °Ré) × 1,875"
                .Item(5).Text = "°Ré × 33/80"
                .Item(6).Text = "°Ré × 21/32 + 7,5"
            End With
        Case Is = 3
            With textHasil
                .Item(0).Text = "K - 273,15"
                .Item(1).Text = "K × 1,8 - 459,67"
                .Item(2).Text = "K × 1,8"
                .Item(3).Text = "(373,15 - K) × 1,5"
                .Item(4).Text = "(K - 273,15) × 33/100"
                .Item(5).Text = "(K - 273,15) × 0,8"
                .Item(6).Text = "(K - 273,15) × 21/40 + 7,5"
            End With
        Case Is = 4
            With textHasil
                .Item(0).Text = "°Ra / 1,8"
                .Item(1).Text = "°Ra / 1,8 + 273,15"
                .Item(2).Text = "°Ra - 459,67"
                .Item(3).Text = "(671,67 - °Ra) × 5/6"
                .Item(4).Text = "(°Ra - 491,67) × 11/60"
                .Item(5).Text = "(°Ra / 1,8 + 273,15) × 0,8"
                .Item(6).Text = "(°Ra - 491,67) × 7/24 + 7,5"
            End With
        Case Is = 5
            With textHasil
                .Item(0).Text = "373,15 - °De × 2/3"
                .Item(1).Text = "100 - °De × 2/3"
                .Item(2).Text = "212 - °De × 1,2"
                .Item(3).Text = "671,67 - °De × 1,2"
                .Item(4).Text = "33 - °De × 0,22"
                .Item(5).Text = "80 - °De × 8/15"
                .Item(6).Text = "60 - °De × 0,35"
            End With
        Case Is = 6
            With textHasil
                .Item(0).Text = "°N × 100/33 + 273,15"
                .Item(1).Text = "°N × 100/33"
                .Item(2).Text = "°N x 60/11 + 32"
                .Item(3).Text = "°N × 60/11 + 491,67"
                .Item(4).Text = "(33 - °N) × 50/11"
                .Item(5).Text = "°N × 80/33"
                .Item(6).Text = "°N × 35/22 + 7,5"
            End With
        Case Is = 7
            With textHasil
                .Item(0).Text = "(°Rø - 7,5) × 40/21 + 273.15"
                .Item(1).Text = "(°Rø - 7,5) × 40/21"
                .Item(2).Text = "(°Rø - 7,5) × 24/7 + 32"
                .Item(3).Text = "(°Rø - 7,5) × 24/7 + 491,67"
                .Item(4).Text = "(60 - °Rø) × 20/7"
                .Item(5).Text = "(°Rø - 7,5) × 22/35"
                .Item(6).Text = "(°Rø - 7,5) × 32/21"
            End With
    End Select
    KunciInput
    cmProses.Enabled = False
End Sub

Private Sub Form_Load()
     AturKontrol
End Sub

Private Sub formCatatanHasil_Click()
    FormCatatanProses.Show vbModal, Me
End Sub

Private Sub menuCelcius_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 1
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuDelisle_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 4
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuFahrenheit_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 2
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuKelvin_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 0
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuNA_Click()
    FormNolAbsolut.Show vbModal, Me
End Sub

Private Sub menuNewton_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 5
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub MenuPengaturan_Click()
    FormPengaturan.Show vbModal, Me
End Sub

Private Sub menuRankine_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 3
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuReamur_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 6
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub

Private Sub menuRomer_Click()
    With FormRingkasan
        .cmbSkala.ListIndex = 7
        .cmbSkala_Click
        .Show vbModal, Me
    End With
End Sub


Private Sub menuTentang_Click()
    MsgBox "Konvertor Suhu " & vbCrLf & _
        "Version: 1.0" & vbCrLf & vbCrLf & _
        "Programmed by Rizky Khapidsyah" & vbCrLf & _
        "All Right Reserved", vbInformation + vbOKOnly, "Tentang"
End Sub

Private Sub textNilaiAwal_KeyPress(KeyAscii As Integer) 'agar karakter yang diisi hanya bisa angka saja
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
    If FormPengaturan.CekDefaultEnter.Value = Checked Then
        If KeyAscii = "13" Then
            cmProses_Click
        End If
    End If
End Sub
