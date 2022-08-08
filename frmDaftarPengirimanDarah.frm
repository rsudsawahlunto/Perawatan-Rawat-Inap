VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPengirimanDarah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pengiriman Darah"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPengirimanDarah.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   14205
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   7440
      Width           =   14175
      Begin VB.CommandButton cmdTerimaDarah 
         Caption         =   "Teri&ma Darah"
         Height          =   465
         Left            =   10200
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   465
         Left            =   12480
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   14175
      Begin VB.Frame Frame3 
         Caption         =   "Periode Kirim Darah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8280
         TabIndex        =   6
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd  MMMM yyyy"
            Format          =   145489923
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd  MMMM yyyy"
            Format          =   145489923
            UpDown          =   -1  'True
            CurrentDate     =   37967
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   7
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12360
      Picture         =   "frmDaftarPengirimanDarah.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPengirimanDarah.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmDaftarPengirimanDarah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCari_Click()
    On Error GoTo errLoad
    Call subLoadData
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTerimaDarah_Click()
    On Error Resume Next
    If dgData.ApproxCount = 0 Then Exit Sub
    With frmTerimaDarahRuangan
        .txtNoOrder = dgData.Columns("NoOrder")
        .txtNoKirim = dgData.Columns("NoKirim")
        .txtNoPendaftaran = dgData.Columns("NoPendaftaran")
        .txtNoCM = dgData.Columns("NoCM")
        .txtNamaPasien = dgData.Columns("NamaPasien")
        .txtJK = dgData.Columns("JK")
        .txtThn = dgData.Columns("UmurTahun")
        .txtBln = dgData.Columns("UmurBulan")
        .txtHr = dgData.Columns("UmurHari")
        .txtSubInstalasi = "-"
        .txtAsalRuangan = dgData.Columns("RuanganPengirim")
        .txtKdRuanganTujuan = dgData.Columns("KdRuangan")
        .fraKirim.Caption = "Terima Darah Dari Ruangan : " & dgData.Columns("RuanganPengirim").Value
        .subLoadData (dgData.Columns("NoKirim"))
        .fraKirim.Enabled = False
        .Show
    End With
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Activate()
    cmdCari_Click
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    Call cmdCari_Click

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub subLoadData()
    On Error GoTo hell_
    Set rs = Nothing
    strSQL = "SELECT TOP (200) NoKirim, TglKirim, NoCM, NoPendaftaran, JmlDarah, NoLabu, TglPeriksa, RuanganPengirim, NamaPasien, NamaDokter, JK, UmurTahun, " & _
    " UmurBulan , UmurHari, RuanganPenerima, NoOrder,KdRuangan, KdRuanganPenerima" & _
    " FROM  V_DaftarPengirimanDarah" & _
    " WHERE TglKirim Between '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
    " and KdRuanganPenerima ='" & mstrKdRuangan & "'"
    Call msubRecFO(rs, strSQL)
    Set dgData.DataSource = rs
    With dgData
        .Columns("TglKirim").Width = 2200
        .Columns("NoKirim").Width = 1200
        .Columns("RuanganPengirim").Width = 2000
        .Columns("NamaPasien").Width = 1500
        .Columns("NoPendaftaran").Width = 2600
        .Columns("NoCM").Width = 1900
        .Columns("NoLabu").Width = 2000
        .Columns("JmlDarah").Width = 1500
        .Columns("NoPendaftaran").Width = 1500
        .Columns("NamaDokter").Width = 2000
        .Columns("JK").Width = 500
        .Columns("UmurTahun").Width = 500
        .Columns("UmurBulan").Width = 500
        .Columns("UmurHari").Width = 500
        .Columns("NoOrder").Width = 1000

        .Columns("KdRuangan").Width = 0
        .Columns("KdRuanganPenerima").Width = 0
        .Columns("UmurBulan").Width = 0
        .Columns("UmurHari").Width = 0
        .Columns("NoOrder").Width = 1000

    End With

    Exit Sub
hell_:
    msubPesanError
End Sub

