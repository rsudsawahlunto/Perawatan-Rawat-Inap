VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Begin VB.Form frmPesanMenuDiet3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Pesan Menu Diet Pasien"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesanMenuDiet3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   11490
   Begin VB.Frame Frame4 
      Caption         =   "Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   26
      Top             =   3360
      Width           =   11175
      Begin VB.Frame Frame5 
         Caption         =   "Umur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   5040
         TabIndex        =   34
         Top             =   840
         Width           =   2415
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   900
            MaxLength       =   6
            TabIndex        =   36
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   6
            TabIndex        =   35
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2130
            TabIndex        =   40
            Top             =   270
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1350
            TabIndex        =   39
            Top             =   277
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   550
            TabIndex        =   38
            Top             =   277
            Width           =   285
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         MaxLength       =   10
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         TabIndex        =   31
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9120
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   9120
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   46
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   45
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9120
         TabIndex        =   44
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   43
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   42
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pendaftaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9120
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NamaPasien"
      Height          =   2175
      Left            =   0
      TabIndex        =   24
      Top             =   1080
      Width           =   11175
      Begin MSDataGridLib.DataGrid dgDaftarPasienRI 
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid dgdiet 
      Height          =   1815
      Left            =   240
      TabIndex        =   23
      Top             =   6840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgUserPemesan 
      Height          =   2055
      Left            =   -3120
      TabIndex        =   15
      Top             =   8520
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   0
      TabIndex        =   19
      Top             =   6600
      Width           =   11175
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   735
         Left            =   9840
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   735
         Left            =   9840
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid fgDiet 
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1720
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Appearance      =   0
      End
   End
   Begin VB.TextBox txtNoOrder 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesan Menu Diet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   4920
      Width           =   11175
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   9840
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   9840
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtUserPemesan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5760
         TabIndex        =   1
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox chkUserPemesan 
         Caption         =   "Pemesan"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcJenisDiet 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisWaktu 
         Height          =   330
         Left            =   6840
         TabIndex        =   4
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   600
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
         Format          =   55640067
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin MSDataListLib.DataCombo dcDetailDiet 
         Height          =   330
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcKeterangan 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detail Diet"
         Height          =   210
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pesan"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Waktu"
         Height          =   210
         Left            =   6960
         TabIndex        =   12
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diet"
         Height          =   210
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   9405
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9360
      Picture         =   "frmPesanMenuDiet3.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPesanMenuDiet3.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmPesanMenuDiet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bolTampilGrid As Boolean

Private Sub loadPesanan()
    Set rs = Nothing
    'strSQL = "SELECT TglOrder, JenisWaktu,  NamaDiet, Keterangan FROM V_PesanMenuDiet WHERE NoCM = '" & mstrNoCM & "'"
    strSQL = "SELECT TglOrder, JenisWaktu,  NamaDiet, Keterangan FROM V_PesanMenuDiet WHERE NoCM = '" & txtNoCM.Text & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgdiet.DataSource = rs
    With dgdiet
    .Columns("0").Width = 2000
    .Columns("1").Width = 1000
    .Columns("2").Width = 2500
    .Columns("3").Width = 3400

    End With
End Sub
Private Sub dgDaftarPasienRI_Click()
    bolTampilGrid = True
    
    '    With frmPesanMenuDiet2
'        .Show
        txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        txtNoCM.Text = dgDaftarPasienRI.Columns(1).Value
        txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            txtSex.Text = "Perempuan"
        Else
            txtSex.Text = "Laki-Laki"
        End If
        txtKls.Text = dgDaftarPasienRI.Columns("Kelas").Value
        txtThn.Text = dgDaftarPasienRI.Columns(11).Value
        txtBln.Text = dgDaftarPasienRI.Columns(12).Value
        txtHr.Text = dgDaftarPasienRI.Columns(13).Value
        txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
        txtTglDaftar.Text = dgDaftarPasienRI.Columns(7).Value
         mdTglMasuk = dgDaftarPasienRI.Columns(7).Value
         mstrKdKelas = dgDaftarPasienRI.Columns(15).Value
         strNoPakai = dgDaftarPasienRI.Columns(10).Value
         mstrKdSubInstalasi = dgDaftarPasienRI.Columns(14)
         mstrNoCM = txtNoCM
         Call loadPesanan
'        End With
'    ElseIf optPasNonAktif.Value = True Then
'        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
'            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
'            Me.Enabled = True
'            Exit Sub
'        End If
'        With frmPesanMenuDiet2
'            .Show
'            .txtNoPendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
'            .txtNoCM.Text = mstrNoCM
'            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
'            .txtSex.Text = dgDaftarPasienRI.Columns(3).Value
'            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
'            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
'            .txtHr.Text = dgDaftarPasienRI.Columns(11).Value
'            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
'            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(12).Value
'            mdTglMasuk = dgDaftarPasienRI.Columns(12).Value
'            mstrKdKelas = dgDaftarPasienRI.Columns(14).Value
'            mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
'        End With

    
    
    
    
    
End Sub
'untuk set grid pasien ri non aktif
Sub SetGridPasienRINonAktif()
    With dgDaftarPasienRI
        .AllowRowSizing = False 'new
        .Columns(0).Width = 1250
        .Columns(0).Caption = "No. Registrasi"
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 750
        .Columns(1).Caption = "No. CM"
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2000
        .Columns(3).Width = 400
        .Columns(4).Width = 1500
        .Columns(5).Width = 1600
        .Columns(6).Width = 1600
        .Columns(6).Caption = "Jenis Pasien"
        .Columns(7).Width = 1590
        .Columns(7).Caption = "Tgl. Pindah"
        .Columns(8).Width = 1750
        .Columns(9).Width = 1650
        .Columns(10).Width = 0 'Umur Tahun
        .Columns(11).Width = 0 'Umur Bulan
        .Columns(12).Width = 0 'Umur Hari
        .Columns(13).Width = 0 'KdSubInstalasi
        .Columns(14).Width = 0 'KdKelas
        .Columns(15).Width = 0 'KdRuangan Tujuan
        .Columns(16).Width = 0 'KdRuangan Asal
        'Columns(17).Width = 0 'KdDiagnosa
    End With
End Sub

Private Sub subSetGrid()
    With fgDiet
        .Cols = 9
        .Rows = 2
        
        .TextMatrix(0, 0) = "Jenis Diet"
        .TextMatrix(0, 1) = "Detail Diet"
        .TextMatrix(0, 2) = "Waktu Diet"
        .TextMatrix(0, 3) = "Keterangan"
        .TextMatrix(0, 4) = "No Order"
        .TextMatrix(0, 5) = "KdJenisDiet"
        .TextMatrix(0, 6) = "KdDiet"
        .TextMatrix(0, 7) = "KdKeterangan"
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 3700
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    
    End With
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
    strSQL = "select kdJenisDiet, jenisDiet from JenisDiet order by JenisDiet"
    Call msubDcSource(dcJenisDiet, rs, strSQL)
    
    strSQL = "select kdJenisWaktu, jenisWaktu from JenisWaktu order by JenisWaktu"
    Call msubDcSource(dcJenisWaktu, rs, strSQL)
    
    strSQL = "select KdDiet, keterangan from Diet where kdDiet = '" & dcDetailDiet.BoundText & "'"
    Call msubDcSource(DcKeterangan, rs, strSQL)
    
    strSQL = "select KdKeterangan, Keterangan from KeteranganMenuDiet order by Keterangan"
    Call msubDcSource(DcKeterangan, rs, strSQL)



Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkUserPemesan_Click()
On Error GoTo errLoad

    If chkUserPemesan.Value = 0 Then
        txtUserPemesan.Enabled = False
        txtUserPemesan.Text = ""
        If dgUserPemesan.Visible = True Then dgUserPemesan.Visible = False
    Else
        txtUserPemesan.Enabled = True
    
        strSQL = "SELECT IdPegawai, NamaLengkap FROM V_DataPegawai"
        Call msubRecFO(rs, strSQL)
    End If

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If chkUserPemesan.Value = 0 Then
    dcJenisDiet.SetFocus
    Else
    txtUserPemesan.SetFocus
    End If
    End If
    
End Sub

Private Sub cmdBatal_Click()
    txtNoOrder.Text = ""
    dtpTglOrder = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""
    dcJenisDiet.Text = ""
    dcDetailDiet.Text = ""
    dcJenisWaktu.Text = ""
'    txtJmlKirim.Text = ""
    dgUserPemesan.Visible = False
End Sub

Private Sub cmdSimpan_Click()
Dim i As Integer
If fgDiet.TextMatrix(1, 0) = "" Then MsgBox "Data Kosong", vbInformation, "Validasi"
If SimpanStrukOrder() = False Then Exit Sub
If SimpanPesanMenuDiet() = False Then Exit Sub

Call SimpanStrukOrder
Call SimpanPesanMenuDiet
    With fgDiet
        For i = 1 To .Rows - 2
        If SimpanDetailPesanMenuDiet(txtNoOrder, .TextMatrix(i, 6), .TextMatrix(i, 7), "A") = False Then Exit Sub
        Next i
    End With

Call Add_HistoryLoginActivity("Add_StrukOrder+Add_PesanMenuDietPasien+Add_DetailPesanMenuDietPasien")
Call loadPesanan
cmdTutup.SetFocus

End Sub

Private Sub cmdTambah_Click()
On Error GoTo errLoad
    
    If Periksa("datacombo", dcJenisDiet, "Jenis diet kosong") = False Then Exit Sub
    If Periksa("datacombo", dcDetailDiet, "Detail diet kosong") = False Then Exit Sub
    If Periksa("datacombo", dcJenisWaktu, "Waktu diet kosong") = False Then Exit Sub
    If Periksa("text", txtUserPemesan, "Pemesan Harus di isi") = False Then Exit Sub
    With fgDiet
        .TextMatrix(.Rows - 1, 0) = dcJenisDiet.Text
        .TextMatrix(.Rows - 1, 1) = dcDetailDiet.Text
        .TextMatrix(.Rows - 1, 2) = dcJenisWaktu.Text
        .TextMatrix(.Rows - 1, 3) = DcKeterangan.Text
        .TextMatrix(.Rows - 1, 4) = txtNoOrder.Text
        .TextMatrix(.Rows - 1, 5) = dcJenisDiet.BoundText
        .TextMatrix(.Rows - 1, 6) = dcDetailDiet.BoundText
        .TextMatrix(.Rows - 1, 7) = DcKeterangan.BoundText
        
        .Rows = .Rows + 1
    End With
    txtNoOrder.Text = ""
    dtpTglOrder = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""
    dcJenisDiet.Text = ""
    dcDetailDiet.Text = ""
    DcKeterangan.BoundText = ""
   
    dgUserPemesan.Visible = False
dtpTglOrder.SetFocus
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Cmdtambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
   frmDaftarPasienRI.Enabled = True
End Sub

Private Sub dcDetailDiet_Change()
   If dcDetailDiet.Text = "" Then Exit Sub
    strSQL = "select KdDiet, keterangan from Diet where kdDiet = '" & dcDetailDiet.BoundText & "'"
'    Call msubDcSource(DcKeterangan, rs, strSQL)
End Sub

Private Sub dcDetailDiet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DcKeterangan.SetFocus
End Sub

Private Sub dcJenisDiet_Change()
On Error GoTo errLoad
    If dcJenisDiet.Text = "" Then Exit Sub
    strSQL = "select KdDiet, NamaDiet, KdJenisDiet from Diet where KdJenisDiet = '" & dcJenisDiet.BoundText & "' ORDER BY NamaDiet"
    Call msubDcSource(dcDetailDiet, rs, strSQL)
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisDiet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcDetailDiet.SetFocus
End Sub


Private Sub dcJenisWaktu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub


Private Sub DcKeterangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dcJenisWaktu.SetFocus
End Sub

Private Sub dgUserPemesan_DblClick()
    Call dgUserPemesan_KeyPress(13)
End Sub

Private Sub dgUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If mintJmlUserPemesan = 0 Then Exit Sub
        txtUserPemesan.Text = dgUserPemesan.Columns(1).Value
        mstrIdPegawai = dgUserPemesan.Columns(0).Value
        If mstrIdPegawai = "" Then
            MsgBox "Pilih dulu User Pemesan yang menangani Pasien", vbCritical, "Validasi"
            txtUserPemesan.Text = ""
            dgUserPemesan.SetFocus
            Exit Sub
        End If
        chkUserPemesan.Value = 1
        dgUserPemesan.Visible = False
        dcJenisDiet.SetFocus
    End If
    If KeyAscii = 27 Then
        dgUserPemesan.Visible = False
    End If
End Sub

Private Sub dtpTglOrder_Change()
    dtpTglOrder.MaxDate = Now
End Sub


Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisDiet.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call openConnection
    dtpTglOrder.Value = Now
    dgUserPemesan.Left = 2400
    dgUserPemesan.Top = 3000
    Call loadPesanan
    Call subSetGrid
    Call subLoadDcSource
    Call cmdBatal_Click
'New
     Set rs = Nothing
     strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    'strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglPindah,[Ruangan Asal],[Ruangan Tujuan],UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,KdRuanganTujuan,KdRuangan from V_DaftarPasienRIPindahKamar where ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%') AND KdRuanganTujuan='" & mstrKdRuangan & "' AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRI.DataSource = rs
    Call SetGridPasienRINonAktif

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRI.Enabled = True
End Sub

Private Sub txtJmlPesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub txtUserPemesan_Change()
    Call subLoadUserPemesan
End Sub

'untuk meload data User Pemesan di grid
Private Sub subLoadUserPemesan()
    strSQL = "SELECT IdPegawai AS [Id Pegawai], NamaLengkap AS [Nama User Pemesan], JenisPegawai FROM V_DataPegawai     mstrFilterUserPemesan WHERE NamaLengkap like '%" & txtUserPemesan.Text & "%'"
    dgUserPemesan.Visible = True

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgUserPemesan
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 3000
       ' .Columns(2).Width = 0
   
          
     
    End With
    dgUserPemesan.Left = 2400
    dgUserPemesan.Top = 2950
End Sub

Private Sub txtUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUserPemesan.Text = "" Then
            MsgBox "Isi dulu User Pemesannya.", vbExclamation, "Validasi"
            txtUserPemesan.SetFocus
        Else
            dgUserPemesan.SetFocus
        End If
    ElseIf KeyAscii = 27 Then
        dgUserPemesan.Visible = False
    End If
End Sub

Private Function SimpanDetailPesanMenuDiet(F_NoOrder As String, F_KdDiet As String, F_Keterangan As String, f_Status As String) As Boolean
SimpanDetailPesanMenuDiet = True
        '================================
        'Simpan Detail Pesan Menu Diet
        '================================
Dim i As Integer
    Set dbcmd = New ADODB.Command
    With dbcmd
'        With fgDiet
'            For i = 1 To .Row - 1
                .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, F_NoOrder)
                .Parameters.Append .CreateParameter("KdDiet", adChar, adParamInput, 6, F_KdDiet)
                .Parameters.Append .CreateParameter("KdKeterangan", adChar, adParamInput, 2, F_Keterangan)
                .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_DetailPesanMenuDietPasien"
                .CommandType = adCmdStoredProc
                .Execute
                    
                If Not (.Parameters("return_value").Value = 0) Then
                    SimpanDetailPesanMenuDiet = False
                    MsgBox "Ada kesalahan dalam pemasukan data Detail Struk Pesan", vbExclamation, "Validasi"
                Else
                  '  MsgBox "Data berhasil disimpan", vbExclamation, "Validasi"
                  
                End If
                Call deleteADOCommandParameters(dbcmd)
'            Next i
'        End With
    End With
End Function
Private Function SimpanStrukOrder() As Boolean

         '====================================
         'simpan Struk Order
         '====================================
      SimpanStrukOrder = True
            Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglOrder.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
       ' .Parameters.Append .CreateParameter("NoOrderGudang", adChar, adParamInput, 20, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, mstrIdPegawai)
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
    
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk order", vbCritical, "Validasi"
            SimpanStrukOrder = False
        Else
            txtNoOrder.Text = .Parameters("OutKode").Value
            Call Add_HistoryLoginActivity("Add_StrukOrder")
        End If
    End With


End Function
 Private Function SimpanPesanMenuDiet() As Boolean
 SimpanPesanMenuDiet = True
  '====================================
         'simpan Pesan menu Diet
         '====================================
         Dim i As Integer
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder)
                .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
                .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
                .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, strNoPakai)
                .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
                .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 6, mstrNoCM)
                .Parameters.Append .CreateParameter("KdJenisWaktu", adChar, adParamInput, 3, dcJenisWaktu.BoundText)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
            
                .ActiveConnection = dbConn
                .CommandText = "dbo.Add_PesanMenuDietPasien"
                .CommandType = adCmdStoredProc
                .Execute
                
                If Not (.Parameters("return_value").Value = 0) Then
                   SimpanPesanMenuDiet = False
                    MsgBox "Ada kesalahan dalam pemasukan data Detail Pesan Menu Diet", vbExclamation, "Validasi"
                Else
                    'MsgBox "Data berhasil disimpan", vbExclamation, "Validasi"
                    
                End If
                Call deleteADOCommandParameters(dbcmd)
            End With
 End Function

