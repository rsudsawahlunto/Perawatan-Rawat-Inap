VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesanMenuDiet2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Pesan Menu Diet Pasien"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesanMenuDiet2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   11220
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
      TabIndex        =   25
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
         TabIndex        =   33
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
         TabIndex        =   32
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         TabIndex        =   30
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8760
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtKls 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   28
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtJenisPasien 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtTglDaftar 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8760
         TabIndex        =   26
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         Left            =   8760
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         Left            =   8760
         TabIndex        =   40
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NamaPasien"
      Height          =   1935
      Left            =   0
      TabIndex        =   24
      Top             =   1080
      Width           =   11175
      Begin VB.CheckBox chkHapusAll 
         Caption         =   "Hapus Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   50
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Pilih Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkCheck 
         Height          =   210
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSFlexGridLib.MSFlexGrid hgDaftarPasienRI 
         Height          =   1575
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2778
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid dgdiet 
      Height          =   1815
      Left            =   240
      TabIndex        =   23
      Top             =   7920
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
      Left            =   120
      TabIndex        =   15
      Top             =   8640
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
      Height          =   3735
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
         Format          =   496435203
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
      Top             =   9300
      Width           =   11220
      _ExtentX        =   19791
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
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9360
      Picture         =   "frmPesanMenuDiet2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPesanMenuDiet2.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmPesanMenuDiet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolTampilGrid As Boolean
Dim rsQuery As New ADODB.recordset
Dim kk As String
Dim ss As Integer

Private Sub loadPesanan()
    Set rs = Nothing
    strSQL = "SELECT TglOrder, JenisWaktu,  NamaDiet, Keterangan FROM V_PesanMenuDiet WHERE NoCM = '" & mstrNoCM & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgdiet.DataSource = rs

    With dgdiet
        .Columns("0").Width = 2000
        .Columns("1").Width = 1000
        .Columns("2").Width = 2500
        .Columns("3").Width = 3400
    End With
End Sub

Private Sub loadPesanan2()
    Dim i As Integer

    i = hgDaftarPasienRI.Row

    Set rs = Nothing
    strSQL = "SELECT TglOrder, JenisWaktu,  NamaDiet, Keterangan FROM V_PesanMenuDiet WHERE NoCM = '" & hgDaftarPasienRI.TextMatrix(i, 3) & "'"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgdiet.DataSource = rs
    With dgdiet
        .Columns("0").Width = 2000
        .Columns("1").Width = 1000
        .Columns("2").Width = 2500
        .Columns("3").Width = 3400

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
    Call msubDcSource(dcJenisdiet, rs, strSQL)

    strSQL = "select kdJenisWaktu, jenisWaktu from JenisWaktu order by JenisWaktu"
    Call msubDcSource(DcJenisWaktu, rs, strSQL)

    strSQL = "select KdDiet, keterangan from Diet where kdDiet = '" & dcDetailDiet.BoundText & "'"
    Call msubDcSource(DcKeterangan, rs, strSQL)

    strSQL = "select KdKeterangan, Keterangan from KeteranganMenuDiet order by Keterangan"
    Call msubDcSource(DcKeterangan, rs, strSQL)

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkAll_Click()

    Call setinputangrid2
    chkAll.Value = 0
End Sub

Private Sub chkHapusAll_Click()
    Call setinputangrid
    chkHapusAll.Value = 0
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
            dcJenisdiet.SetFocus
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
    dcJenisdiet.Text = ""
    dcDetailDiet.Text = ""
    DcJenisWaktu.Text = ""
    dgUserPemesan.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    Dim j As Integer

    If fgDiet.TextMatrix(1, 0) = "" Then MsgBox "Data Kosong", vbInformation, "Validasi"

    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic

    For i = 1 To rs.RecordCount
        If hgDaftarPasienRI.TextMatrix(i, 1) <> Chr$(187) Then

        End If
        If hgDaftarPasienRI.TextMatrix(i, 1) = Chr$(187) Then

            ss = i
            txtNoOrder.Text = ""
            If SimpanStrukOrder() = False Then Exit Sub
            If SimpanPesanMenuDiet() = False Then Exit Sub

            Call SimpanStrukOrder
            Call SimpanPesanMenuDiet

            With fgDiet
                For j = 1 To .Rows - 2
                    If SimpanDetailPesanMenuDiet(txtNoOrder, .TextMatrix(j, 6), .TextMatrix(j, 7), "A") = False Then Exit Sub

                Next j
            End With

        End If

        Call loadPesanan2

        rs.MoveNext

    Next i

    Call loadPesanan2
    Call ClearFgDiet

    Call Add_HistoryLoginActivity("Add_StrukOrder+Add_PesanMenuDietPasien+Add_DetailPesanMenuDietPasien")
    Call loadPesanan
    cmdTutup.SetFocus

End Sub

Private Sub cmdTambah_Click()
    On Error GoTo errLoad

    If Periksa("datacombo", dcJenisdiet, "Jenis diet kosong") = False Then Exit Sub
    If Periksa("datacombo", dcDetailDiet, "Detail diet kosong") = False Then Exit Sub
    If Periksa("datacombo", DcJenisWaktu, "Waktu diet kosong") = False Then Exit Sub
    If Periksa("text", txtUserPemesan, "Pemesan Harus di isi") = False Then Exit Sub
    With fgDiet
        .TextMatrix(.Rows - 1, 0) = dcJenisdiet.Text
        .TextMatrix(.Rows - 1, 1) = dcDetailDiet.Text
        .TextMatrix(.Rows - 1, 2) = DcJenisWaktu.Text
        .TextMatrix(.Rows - 1, 3) = DcKeterangan.Text
        .TextMatrix(.Rows - 1, 4) = txtNoOrder.Text
        .TextMatrix(.Rows - 1, 5) = dcJenisdiet.BoundText
        .TextMatrix(.Rows - 1, 6) = dcDetailDiet.BoundText
        .TextMatrix(.Rows - 1, 7) = DcKeterangan.BoundText

        .Rows = .Rows + 1
    End With
    txtNoOrder.Text = ""
    dtpTglOrder = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""

    dgUserPemesan.Visible = False
    dtpTglOrder.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub ClearFgDiet()
    On Error GoTo errLoad

    With fgDiet
        .TextMatrix(.Rows - 1, 0) = ""
        .TextMatrix(.Rows - 1, 1) = ""
        .TextMatrix(.Rows - 1, 2) = ""
        .TextMatrix(.Rows - 1, 3) = ""
        .TextMatrix(.Rows - 1, 4) = ""
        .TextMatrix(.Rows - 1, 5) = ""
        .TextMatrix(.Rows - 1, 6) = ""
        .TextMatrix(.Rows - 1, 7) = ""

        .Rows = .Rows + 1
    End With
    txtNoOrder.Text = ""
    dtpTglOrder = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""

    dgUserPemesan.Visible = False
    dtpTglOrder.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    frmDaftarPasienRI.Enabled = True
End Sub

Private Sub dcDetailDiet_Change()
    If dcDetailDiet.Text = "" Then Exit Sub
    strSQL = "select KdDiet, keterangan from Diet where kdDiet = '" & dcDetailDiet.BoundText & "'"
End Sub

Private Sub dcDetailDiet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DcKeterangan.SetFocus
End Sub

Private Sub dcJenisDiet_Change()
    On Error GoTo errLoad
    If dcJenisdiet.Text = "" Then Exit Sub
    strSQL = "select KdDiet, NamaDiet, KdJenisDiet from Diet where KdJenisDiet = '" & dcJenisdiet.BoundText & "' ORDER BY NamaDiet"
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
    If KeyAscii = 13 Then DcJenisWaktu.SetFocus
End Sub

Private Sub dgUserPemesan_DblClick()
    Call dgUserPemesan_KeyPress(13)
End Sub

Private Sub dgUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

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
        dcJenisdiet.SetFocus
    End If
    If KeyAscii = 27 Then
        dgUserPemesan.Visible = False
    End If
End Sub

Private Sub dtpTglOrder_Change()
    dtpTglOrder.MaxDate = Now
End Sub

Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisdiet.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call openConnection
    dtpTglOrder.Value = Now
    dgUserPemesan.Left = 2400
    dgUserPemesan.Top = 5000
    Call loadPesanan
    Call subSetGrid
    Call subLoadDcSource
    Call cmdBatal_Click
    Call setClearGridTagihan

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRI.Enabled = True
End Sub

Private Sub txtJmlPesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub hgDaftarPasienRI_Click()
    Dim i As Integer

    Set rs = Nothing
    strQuery = "select '',NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic

    Call loadPesanan2

    i = hgDaftarPasienRI.Row
    txtNoPendaftaran.Text = hgDaftarPasienRI.TextMatrix(i, 2)
    txtNoCM.Text = hgDaftarPasienRI.TextMatrix(i, 3)
    txtNamaPasien.Text = hgDaftarPasienRI.TextMatrix(i, 4)

    If hgDaftarPasienRI.TextMatrix(i, 5) = "P" Then
        txtSex.Text = "Perempuan"
    Else
        txtSex.Text = "Laki-Laki"
    End If

    txtKls.Text = hgDaftarPasienRI.TextMatrix(i, 7)
    txtThn.Text = hgDaftarPasienRI.TextMatrix(i, 13)
    txtBln.Text = hgDaftarPasienRI.TextMatrix(i, 14)
    txtHr.Text = hgDaftarPasienRI.TextMatrix(i, 15)
    txtJenisPasien.Text = hgDaftarPasienRI.TextMatrix(i, 8)
    txtTglDaftar.Text = hgDaftarPasienRI.TextMatrix(i, 9)

End Sub

Private Sub hgDaftarPasienRI_DblClick()

    If hgDaftarPasienRI.Rows = 1 Then Exit Sub
    If hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, 3) = "" Then Exit Sub
    chkCheck.Visible = False

    Select Case hgDaftarPasienRI.Col
        Case 1
            chkCheck.Top = hgDaftarPasienRI.RowPos(hgDaftarPasienRI.Row) + 280
            Dim intA As Integer
            chkCheck.Visible = True
            intA = ((hgDaftarPasienRI.ColPos(hgDaftarPasienRI.Col + 1) - hgDaftarPasienRI.ColPos(hgDaftarPasienRI.Col)) / 2)
            chkCheck.Left = hgDaftarPasienRI.ColPos(hgDaftarPasienRI.Col) + 50 + intA
            chkCheck.SetFocus
            If hgDaftarPasienRI.Col = 1 Then
                If hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, 1) <> "" Then
                    chkCheck.Value = 1
                Else
                    chkCheck.Value = 0
                End If
            End If
    End Select

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
    End With
    dgUserPemesan.Left = 2400
    dgUserPemesan.Top = 5000
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

Private Function SimpanDetailPesanMenuDiet(F_NoOrder As String, F_KdDiet As String, f_Keterangan As String, f_status As String) As Boolean
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
        .Parameters.Append .CreateParameter("KdDiet", adChar, adParamInput, 6, dcDetailDiet.BoundText)
        .Parameters.Append .CreateParameter("KdKeterangan", adChar, adParamInput, 2, DcKeterangan.BoundText)
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailPesanMenuDietPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            SimpanDetailPesanMenuDiet = False
            MsgBox "Ada kesalahan dalam pemasukan data Detail Struk Pesan", vbExclamation, "Validasi"
        Else

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
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, mstrIdPegawai)
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

    Dim i As Integer
    Dim j As Integer
    SimpanPesanMenuDiet = True

    '====================================
    'simpan Pesan menu Diet
    '====================================

    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, hgDaftarPasienRI.TextMatrix(ss, 12))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, hgDaftarPasienRI.TextMatrix(ss, 2))
        .Parameters.Append .CreateParameter("NoCM", adChar, adParamInput, 10, hgDaftarPasienRI.TextMatrix(ss, 3))
        .Parameters.Append .CreateParameter("KdJenisWaktu", adChar, adParamInput, 3, DcJenisWaktu.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PesanMenuDietPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            SimpanPesanMenuDiet = False
            MsgBox "Ada kesalahan dalam pemasukan data Detail Pesan Menu Diet", vbExclamation, "Validasi"
        Else

        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

'Tahap 1 => Membuat size grid
Private Sub setClearGridTagihan()
    Dim i As Integer
    With hgDaftarPasienRI
        .clear
        .Rows = 2
        .Cols = 20

        'WindRunner
        .ColWidth(0) = 0 '320
        .ColWidth(1) = 200 ' <<<
        .ColWidth(2) = 1200 'NoPendataran
        .ColWidth(3) = 700 'NoCM
        .ColWidth(4) = 1500 'Nama Pasien
        .ColWidth(5) = 300 'JK
        .ColWidth(6) = 1200 'Umur
        .ColWidth(7) = 1200 'Kelas
        .ColWidth(8) = 1200 'Jenis Pasien
        .ColWidth(9) = 1800 'TglPendaftaran
        .ColWidth(10) = 0 'NoKamar
        .ColWidth(11) = 0 'NoBed
        .ColWidth(12) = 0 'NoPakai
        .ColWidth(13) = 600 'Tahun
        .ColWidth(14) = 500 'Bulan
        .ColWidth(15) = 500 'Hari
        .ColWidth(16) = 0 'KdSubInstalasi
        .ColWidth(17) = 0 'KdKelas
        .ColWidth(18) = 0 'CaraMasuk
        .ColWidth(19) = 0 '

        Call setJudulTagihan
        'Call txtNoPendaftaran_KeyPress(13)
        Call setinputangrid
    End With
End Sub

'Tahap 2 => Penamaan colum grid
Private Sub setJudulTagihan()
    Dim i As Integer

    With hgDaftarPasienRI
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "NoPendaftaran"
        .TextMatrix(0, 3) = "NoCM"
        .TextMatrix(0, 4) = "Nama Pasien"
        .TextMatrix(0, 5) = "JK"
        .TextMatrix(0, 6) = "Umur"
        .TextMatrix(0, 7) = "Kelas"
        .TextMatrix(0, 8) = "JenisPasien"
        .TextMatrix(0, 9) = "TglPendaftaran"
        .TextMatrix(0, 10) = "NoKamar"
        .TextMatrix(0, 11) = "NoBed"
        .TextMatrix(0, 12) = "NoPakai"
        .TextMatrix(0, 13) = "Tahun"
        .TextMatrix(0, 14) = "Bulan"
        .TextMatrix(0, 15) = "Hari"
        .TextMatrix(0, 16) = "KdSubInstalasi"
        .TextMatrix(0, 17) = "KdKelas"
        .TextMatrix(0, 18) = "CaraMasuk"
    End With
End Sub

'Tahap 3 => Penginputan data ke grid
Public Sub setinputangrid()
    On Error GoTo errLoad
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    chkCheck.Top = 500
    chkCheck.Left = 150

    Set rs = Nothing
    strQuery = "select '', NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic

    If rs.RecordCount <> 0 Then
        hgDaftarPasienRI.clear
        hgDaftarPasienRI.Rows = rs.RecordCount + 1

        For i = 1 To rs.RecordCount
            For j = 1 To 18

                hgDaftarPasienRI.TextMatrix(i, j) = "" & rs(j - 1).Value

                If j = 1 Then kk = hgDaftarPasienRI.TextMatrix(i, j)
                If j = 1 Then hgDaftarPasienRI.TextMatrix(i, j) = ""
                'Chr$ (187)

            Next j
            rs.MoveNext
        Next i

        Call setJudulTagihan
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Tahap 3 => Penginputan data ke grid
Public Sub setinputangrid2()
    On Error GoTo errLoad
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    chkCheck.Top = 500
    chkCheck.Left = 150

    Set rs = Nothing
    strQuery = "select '', NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic

    If rs.RecordCount <> 0 Then
        hgDaftarPasienRI.clear
        hgDaftarPasienRI.Rows = rs.RecordCount + 1

        For i = 1 To rs.RecordCount
            For j = 1 To 18

                hgDaftarPasienRI.TextMatrix(i, j) = "" & rs(j - 1).Value

                If j = 1 Then kk = hgDaftarPasienRI.TextMatrix(i, j)
                If j = 1 Then hgDaftarPasienRI.TextMatrix(i, j) = Chr$(187)

            Next j
            rs.MoveNext
        Next i

        Call setJudulTagihan
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkCheck_Click()
    On Error GoTo errLoad

    If chkCheck.Value = vbChecked Then
        hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, hgDaftarPasienRI.Col) = Chr$(187)
        hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, 19) = 1
    Else
        hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, hgDaftarPasienRI.Col) = ""
        chkCheck.Visible = False
        hgDaftarPasienRI.TextMatrix(hgDaftarPasienRI.Row, 19) = 0
    End If
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkCheck.Visible = False
        Call chkCheck_Click
        hgDaftarPasienRI.SetFocus
    End If
End Sub
