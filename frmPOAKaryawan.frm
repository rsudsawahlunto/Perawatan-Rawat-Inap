VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPOAKaryawan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Obat & Alkes Karyawan"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOAKaryawan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPOAKaryawan.frx":0CCA
   ScaleHeight     =   6510
   ScaleWidth      =   10965
   Begin VB.TextBox txtAsalBarang 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   28
      Text            =   "txtAsalBarang"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSatuan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   27
      Text            =   "txtSatuan"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdAsal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   26
      Text            =   "txtKdAsal"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdBarang 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Text            =   "txtKdBarang"
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtKdDokter 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Text            =   "txtKdDokter"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   3015
      Left            =   9840
      TabIndex        =   21
      Top             =   -720
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgHargaBrg 
      Height          =   3015
      Left            =   -3120
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid dgData 
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   50
      Cols            =   6
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   8577768
      ForeColorFixed  =   -2147483627
      ForeColorSel    =   -2147483628
      BackColorBkg    =   16777215
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   10935
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   8760
         TabIndex        =   23
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   9735
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtjml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "1"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txthargasatuan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txttotbiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNamaBrg 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   210
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Harga Satuan"
         Height          =   210
         Left            =   5280
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   6960
         TabIndex        =   16
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   10935
      Begin VB.TextBox txtKeperluan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   146079747
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Keperluan Pemakaian"
         Height          =   210
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Penanggung Jawab"
         Height          =   210
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pemakaian"
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1560
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   10935
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9480
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   29
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPOAKaryawan.frx":190C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9120
      Picture         =   "frmPOAKaryawan.frx":42CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPOAKaryawan.frx":57BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPOAKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
