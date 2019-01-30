VERSION 5.00
Object = "{A30DC858-670B-4336-A74E-10C38ADF5ADD}#1.0#0"; "xTab.ocx"
Object = "{530871E2-C21C-4628-9427-E2C09620063B}#1.0#0"; "XP_Engine.ocx"
Object = "{9333B604-42A2-408B-AE2E-ED76202997B7}#2.0#0"; "isButton.ocx"
Begin VB.Form FormPengaturan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pengaturan"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPengaturan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin isButton3.isButton cmOK 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Icon            =   "FormPengaturan.frx":000C
      Style           =   7
      Caption         =   "&OK"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjXTab.XTab XTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      TabCount        =   2
      TabCaption(0)   =   "Umum"
      TabContCtrlCnt(0)=   5
      Tab(0)ContCtrlCap(1)=   "CekDefaultEnter"
      Tab(0)ContCtrlCap(2)=   "cmbDefaultSimpan"
      Tab(0)ContCtrlCap(3)=   "cekSimpanHasilProses"
      Tab(0)ContCtrlCap(4)=   "Label1"
      Tab(0)ContCtrlCap(5)=   "Label2"
      TabCaption(1)   =   "System"
      TabContCtrlCnt(1)=   3
      Tab(1)ContCtrlCap(1)=   "cmbDibelakangKoma"
      Tab(1)ContCtrlCap(2)=   "CekBulankanHasilPecahan"
      Tab(1)ContCtrlCap(3)=   "LabelDiBekalangKoma"
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   14737632
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.ComboBox cmbDibelakangKoma 
         Enabled         =   0   'False
         Height          =   390
         Left            =   -73200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox CekBulankanHasilPecahan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "  Bulatkan Hasil Pecahan"
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox CekDefaultEnter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aktifkan Default Enter"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbDefaultSimpan 
         Height          =   390
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CheckBox cekSimpanHasilProses 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simpan Hasil Proses ke Dalam Catatan"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Label LabelDiBekalangKoma 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Di Belakang koma :"
         Enabled         =   0   'False
         Height          =   270
         Left            =   -74520
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default File Simpan"
         Height          =   270
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   2040
         TabIndex        =   4
         Top             =   2400
         Width           =   45
      End
   End
   Begin XPEngine.XPControl MesinXP 
      Left            =   0
      Top             =   6360
      _ExtentX        =   529
      _ExtentY        =   503
      ColorScheme     =   3
   End
End
Attribute VB_Name = "FormPengaturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    With cmbDefaultSimpan
        .Clear
        .AddItem "RikySoft Catatan File (.rcf)", 0
        .AddItem "Microsoft Word 2003 Document (.doc)", 1
        .AddItem "Microsoft Excel 2003 Document (.xls)", 2
        .AddItem "Microsoft Rich Text Format Document (.rtf)", 3
        .AddItem "Text File (.txt)", 4
        .ListIndex = 0
    End With
    With cmbDibelakangKoma
        .Clear
        .AddItem "1", 0
        .AddItem "2", 1
        .AddItem "3", 2
        .AddItem "4", 3
        .AddItem "5", 4
        .AddItem "6", 5
        .AddItem "7", 6
        .AddItem "8", 7
        .AddItem "9", 8
        .AddItem "10", 9
        .ListIndex = 0
    End With
End Sub

Private Sub CekBulankanHasilPecahan_Click()
    If CekBulankanHasilPecahan.Value = Checked Then
        With Me
            .LabelDiBekalangKoma.Enabled = True
            .cmbDibelakangKoma.Enabled = True
        End With
    ElseIf CekBulankanHasilPecahan.Value = Unchecked Then
        With Me
            .LabelDiBekalangKoma.Enabled = False
            .cmbDibelakangKoma.Enabled = False
        End With
    End If
End Sub

Private Sub cmOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

