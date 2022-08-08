VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmResumePulang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resume Medis Pulang"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResumePulang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   13470
   Begin VB.Frame Frame7 
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   5640
      Width           =   13455
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   11880
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Riwayat Penyakit"
      TabPicture(0)   =   "frmResumePulang.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgRiwayatPenyakit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraRiwayatPenyakit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pemeriksaan Fisik"
      TabPicture(1)   =   "frmResumePulang.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgPemeriksaanFisik"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdTambahPemeriksaanFisik"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraPemeriksaanFisik"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Diagnosa"
      TabPicture(2)   =   "frmResumePulang.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dgDiagnosa"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdHapusDiagnosa"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdTambahDiagnosa"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraCari"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Prosedur Diagnostik"
      TabPicture(3)   =   "frmResumePulang.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Obat"
      TabPicture(4)   =   "frmResumePulang.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dgObat"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraCariObat"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Obat Pulang"
      TabPicture(5)   =   "frmResumePulang.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dgObatPulang"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "fraCariObatPulang"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Instruksi Lanjutan"
      TabPicture(6)   =   "frmResumePulang.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label2"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label3"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label5"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Label6"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "txtKeadaanPulang"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "txtAlasanRawat"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "txtRencana"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "dcKontrolKe"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "dtpTglKontrol"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "cmdSimpanInstruksiLanjutan"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).ControlCount=   11
      Begin VB.Frame fraCariObatPulang 
         Height          =   2775
         Left            =   -68880
         TabIndex        =   61
         Top             =   1680
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox txtCariObatPulang 
            Height          =   375
            Left            =   720
            TabIndex        =   62
            Top             =   240
            Width           =   3735
         End
         Begin MSDataGridLib.DataGrid dgCariObatPulang 
            Height          =   1935
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
         Begin VB.Label Label19 
            Caption         =   "Cari"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraCariObat 
         Height          =   2775
         Left            =   -68880
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox txtCariObat 
            Height          =   375
            Left            =   720
            TabIndex        =   53
            Top             =   240
            Width           =   3735
         End
         Begin MSDataGridLib.DataGrid dgCariObat 
            Height          =   1935
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
         Begin VB.Label Label16 
            Caption         =   "Cari"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3975
         Left            =   -66600
         TabIndex        =   65
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton cmdTambahObatPulang 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            TabIndex        =   73
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdHapusObatPulang 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   72
            Top             =   3480
            Width           =   1095
         End
         Begin VB.TextBox txtObatPulang 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdCariObatPulang 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   66
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "Obat Pulang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3975
         Left            =   -66600
         TabIndex        =   56
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton cmdHapusObat 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   71
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton cmdTambahObat 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            TabIndex        =   70
            Top             =   3480
            Width           =   1095
         End
         Begin VB.TextBox txtObat 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdCariObat 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   57
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Obat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fraCari 
         Height          =   2775
         Left            =   -68880
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox txtCariDiagnosa 
            Height          =   375
            Left            =   720
            TabIndex        =   49
            Top             =   240
            Width           =   3735
         End
         Begin MSDataGridLib.DataGrid dgCariDiagnosa 
            Height          =   1935
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
         Begin VB.Label Label14 
            Caption         =   "Cari"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3495
         Left            =   -66600
         TabIndex        =   40
         Top             =   480
         Width           =   4935
         Begin VB.ComboBox cmbJenisDiagnosa 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmResumePulang.frx":0D8E
            Left            =   120
            List            =   "frmResumePulang.frx":0D98
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   43
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtDiagnosa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label15 
            Caption         =   "Jenis Diagnosa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Diagnosa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdTambahDiagnosa 
         Caption         =   "&Tambah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65280
         TabIndex        =   39
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapusDiagnosa 
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66480
         TabIndex        =   38
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Frame fraPemeriksaanFisik 
         Height          =   3855
         Left            =   -66480
         TabIndex        =   33
         Top             =   600
         Width           =   4815
         Begin VB.CommandButton cmdHapusPemeriksaanFisik 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            TabIndex        =   74
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtPemeriksaanFisik 
            Height          =   1125
            Left            =   360
            TabIndex        =   35
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdSimpanPemeriksaanFisik 
            Caption         =   "&Simpan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   34
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Pemeriksaan Fisik"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraRiwayatPenyakit 
         Height          =   3855
         Left            =   8400
         TabIndex        =   29
         Top             =   570
         Width           =   4935
         Begin VB.CommandButton cmdTambahRiwayatPenyakit 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            TabIndex        =   45
            Top             =   2100
            Width           =   1095
         End
         Begin VB.CommandButton cmdHapusss 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   44
            Top             =   2100
            Width           =   1095
         End
         Begin VB.TextBox txtRiwayatPenyakit 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label Label11 
            Caption         =   "Riwayat Penyakit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -74880
         TabIndex        =   25
         Top             =   3600
         Width           =   13215
         Begin VB.CommandButton cmdSimpanProsedurDiagnostik 
            Caption         =   "&SIMPAN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   12000
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   13215
         Begin VB.TextBox txtLab 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   20
            Top             =   360
            Width           =   9150
         End
         Begin VB.TextBox txtRad 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   19
            Top             =   840
            Width           =   9150
         End
         Begin VB.TextBox txtEKG 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   18
            Top             =   1320
            Width           =   9150
         End
         Begin VB.TextBox txtPD 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   17
            Top             =   1800
            Width           =   9150
         End
         Begin VB.Label Label7 
            Caption         =   "Laboratorium"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Radiologi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "EKG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Prosedur Diagnostik"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1800
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdTambahPemeriksaanFisik 
         Caption         =   "&Tambah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -64320
         TabIndex        =   15
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSimpanInstruksiLanjutan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67320
         TabIndex        =   14
         Top             =   3840
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpTglKontrol 
         Height          =   375
         Left            =   -72720
         TabIndex        =   13
         Top             =   2640
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
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
         Format          =   155385857
         CurrentDate     =   44467
      End
      Begin MSDataListLib.DataCombo dcKontrolKe 
         Height          =   330
         Left            =   -72720
         TabIndex        =   11
         Top             =   2160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
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
      Begin VB.TextBox txtRencana 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72720
         TabIndex        =   8
         Top             =   3120
         Width           =   6630
      End
      Begin VB.TextBox txtAlasanRawat 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72720
         TabIndex        =   6
         Top             =   1680
         Width           =   6630
      End
      Begin VB.TextBox txtKeadaanPulang 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72720
         TabIndex        =   4
         Top             =   1200
         Width           =   6630
      End
      Begin MSDataGridLib.DataGrid dgRiwayatPenyakit 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
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
      Begin MSDataGridLib.DataGrid dgPemeriksaanFisik 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   32
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6588
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
      Begin MSDataGridLib.DataGrid dgDiagnosa 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   37
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6800
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
      Begin MSDataGridLib.DataGrid dgObat 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   60
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6800
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
      Begin MSDataGridLib.DataGrid dgObatPulang 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   69
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6800
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
      Begin VB.Label Label6 
         Caption         =   "Tanggal Kontrol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Kontrol Ke"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Rencana Tindak Lanjut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Alasan Rawat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Keadaan Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin VB.Label Label4 
      Caption         =   "Rencana Tindak Lanjut"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmResumePulang.frx":0DBF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmResumePulang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    If fraCari.Visible = False Then
        fraCari.Visible = True
    Else
        fraCari.Visible = False
    End If
End Sub

Private Sub cmdCariObat_Click()
    If fraCariObat.Visible = False Then
        fraCariObat.Visible = True
    Else
        fraCariObat.Visible = False
    End If
End Sub

Private Sub cmdCariObatPulang_Click()
    If fraCariObatPulang.Visible = False Then
        fraCariObatPulang.Visible = True
    Else
        fraCariObatPulang.Visible = False
    End If
End Sub

Private Sub cmdHapusDiagnosa_Click()
    Dim vbMsgResult As VbMsgBoxResult
    On Error GoTo errLoad
    
    If dgDiagnosa.ApproxCount = 0 Then Exit Sub
    vbMsgResult = MsgBox("Apakah anda yakin ingin menghapus data Diagnosa?", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgResult = vbNo Then Exit Sub
    
    Call sp_HapusDiagnosa(dbcmd)
    Call subLoadDiagnosa
    
    MsgBox "Penghapusan data Pemeriksaan Fisik berhasil", vbInformation, "Informasi"
    Exit Sub
errLoad:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusObat_Click()
    Dim vbMsgResult As VbMsgBoxResult
    On Error GoTo errLoad
    
    If dgObat.ApproxCount = 0 Then Exit Sub
    vbMsgResult = MsgBox("Apakah anda yakin ingin menghapus data Obat?", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgResult = vbNo Then Exit Sub
    
    Call sp_HapusObat(dbcmd)
    Call subLoadObat
    
    MsgBox "Penghapusan data Obat berhasil", vbInformation, "Informasi"
    Exit Sub
errLoad:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusPemeriksaanFisik_Click()
    Dim vbMsgResult As VbMsgBoxResult
    On Error GoTo errLoad
    
    If dgPemeriksaanFisik.ApproxCount = 0 Then Exit Sub
    vbMsgResult = MsgBox("Apakah anda yakin ingin menghapus data Pemeriksaan Fisik?", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgResult = vbNo Then Exit Sub
    
    Call sp_hapusPemeriksaanFisik(dbcmd)
    Call subLoadPemeriksaanFisik
    
    MsgBox "Penghapusan data Pemeriksaan Fisik berhasil", vbInformation, "Informasi"
    Exit Sub
errLoad:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdHapusss_Click()
    Dim vbMsgResult As VbMsgBoxResult
    On Error GoTo errLoad
    
    If dgRiwayatPenyakit.ApproxCount = 0 Then Exit Sub
    vbMsgResult = MsgBox("Apakah anda yakin ingin menghapus data Riwayat Penyakit?", vbQuestion + vbYesNo, "Konfirmasi")
    If vbMsgResult = vbNo Then Exit Sub
    
    Call sp_hapusRiwayatPenyakit(dbcmd)
    Call subLoadRiwayatPenyakit
    
    MsgBox "Penghapusan data Riwayat Penyakit berhasil", vbInformation, "Informasi"
    Exit Sub
errLoad:
    MsgBox "Data gagal dihapus, hubungi administrator", vbCritical, "Validasi"
End Sub

Private Sub cmdKeluar_Click()
    frmTransaksiPasien.Enabled = True
    Unload Me
End Sub

Private Sub cmdSimpanInstruksiLanjutan_Click()
    Call sp_SimpanIL(dbcmd)
    Call subLoadIL
End Sub

Private Sub cmdSimpanPemeriksaanFisik_Click()
    If txtPemeriksaanFisik = "" Then
        MsgBox "Pemeriksaan Fisik belum diisi!!!", vbExclamation, "Peringatan"
        txtPemeriksaanFisik.SetFocus
        Exit Sub
    End If
    
    Call sp_simpanPemeriksaanFisik(dbcmd)
    Call subLoadPemeriksaanFisik
    fraPemeriksaanFisik.Visible = False
End Sub

Private Sub cmdSimpanRiwayatPenyakit_Click()

End Sub

Private Sub cmdSimpanProsedurDiagnostik_Click()
    If txtLab.Text = "" Then
        MsgBox "Data Laboratorium belum diisi!!!", vbExclamation, "Peringatan"
        txtLab.SetFocus
        Exit Sub
    End If
    If txtRad.Text = "" Then
        MsgBox "Data Radiologi belum diisi!!!", vbExclamation, "Peringatan"
        txtRad.SetFocus
        Exit Sub
    End If
    If txtEKG.Text = "" Then
        MsgBox "Data EKG belum diisi!!!", vbExclamation, "Peringatan"
        txtEKG.SetFocus
        Exit Sub
    End If
    If txtPD.Text = "" Then
        MsgBox "Data Pemeriksaan Diagnostik belum diisi!!!", vbExclamation, "Peringatan"
        txtPD.SetFocus
        Exit Sub
    End If
    
    Call sp_SimpanPD(dbcmd)
    Call subLoadPD
End Sub

Private Sub cmdTambahDiagnosa_Click()
    If txtDiagnosa.Text = "" Then
        MsgBox "Data Diagnosa belum diisi!!!", vbExclamation, "Peringatan"
        txtRiwayatPenyakit.SetFocus
        Exit Sub
    End If
    If cmbJenisDiagnosa.Text = "" Then
        MsgBox "Jenis Diagnosa belum diisi!!!", vbExclamation, "Peringatan"
        cmbJenisDiagnosa.SetFocus
        Exit Sub
    End If
    
    Call sp_SimpanDiagnosa(dbcmd)
    Call subLoadDiagnosa
End Sub

Private Sub cmdTambahObat_Click()
    If txtObat.Text = "" Then
        MsgBox "Data Obat belum diisi!!!", vbExclamation, "Peringatan"
        'txtRiwayatPenyakit.SetFocus
        Exit Sub
    End If
    
    Call sp_SimpanObat(dbcmd)
    Call subLoadObat
    txtCariObat.Text = ""
    subLoadCariObat
End Sub

Private Sub cmdTambahObatPulang_Click()
    If txtObatPulang.Text = "" Then
        MsgBox "Data Obat belum diisi!!!", vbExclamation, "Peringatan"
        'txtRiwayatPenyakit.SetFocus
        Exit Sub
    End If
    
    Call sp_SimpanObatPulang(dbcmd)
    Call subLoadObatPulang
    txtCariObatPulang.Text = ""
    subLoadCariObatPulang
End Sub

Private Sub cmdTambahPemeriksaanFisik_Click()
    fraPemeriksaanFisik.Visible = True
End Sub

Private Sub cmdTambahRiwayatPenyakit_Click()
    If txtRiwayatPenyakit = "" Then
        MsgBox "Riwayat Penyakit belum diisi!!!", vbExclamation, "Peringatan"
        txtRiwayatPenyakit.SetFocus
        Exit Sub
    End If
    
    Call sp_simpanRiwayatPenyakit(dbcmd)
    Call subLoadRiwayatPenyakit
    'fraRiwayatPenyakit.Visible = False
End Sub


Private Sub dgCariDiagnosa_DblClick()
    If dgCariDiagnosa.ApproxCount = 0 Then Exit Sub
    txtDiagnosa.Text = dgCariDiagnosa.Columns("Nama Diagnosa").Value
    fraCari.Visible = False
End Sub

Private Sub dgCariObat_DblClick()
    If dgCariObat.ApproxCount = 0 Then Exit Sub
    txtObat.Text = dgCariObat.Columns("Nama Obat").Value
    fraCariObat.Visible = False
End Sub

Private Sub dgCariObatPulang_DblClick()
    If dgCariObatPulang.ApproxCount = 0 Then Exit Sub
    txtObatPulang.Text = dgCariObatPulang.Columns("Nama Obat").Value
    fraCariObatPulang.Visible = False
End Sub

Private Sub Form_Load()
    dtpTglKontrol.Value = Now + 1
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadRiwayatPenyakit
    Call subLoadPemeriksaanFisik
    Call subLoadCariDiagnosa
    Call subLoadDiagnosa
    Call subLoadPD
    Call subLoadObat
    Call subLoadCariObat
    Call subLoadObatPulang
    Call subLoadCariObatPulang
    Call subLoadIL
    
    strSQL = "select KdRuangan, NamaRuangan from Ruangan where KdInstalasi='02' and StatusEnabled=1 and KdRuangan<>'220' order by NamaRuangan"
    Call msubDcSource(dcKontrolKe, rs, strSQL)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    Unload Me
End Sub

Public Sub subLoadRiwayatPenyakit()
    On Error GoTo errLoad
    
    strSQL = "SELECT id,RiwayatPenyakit AS [Riwayat Penyakit] FROM RMPRiwayatPenyakit WHERE " _
    & "NoPendaftaran='" & mstrNoPen & "' ORDER BY id"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgRiwayatPenyakit.DataSource = rs
    With dgRiwayatPenyakit
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With

    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadPemeriksaanFisik()
    On Error GoTo errLoad
    
    strSQL = "SELECT id,PemeriksaanFisik AS [Pemeriksaan Fisik] FROM RMPPemeriksaanFisik WHERE " _
    & "NoPendaftaran='" & mstrNoPen & "' ORDER BY id"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgPemeriksaanFisik.DataSource = rs
    With dgPemeriksaanFisik
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With

    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadCariDiagnosa()
    On Error GoTo errLoad
    
    strSQL = "SELECT KdDiagnosa AS Kode, NamaDiagnosa AS [Nama Diagnosa] FROM Diagnosa WHERE " _
    & "(KdDiagnosa LIKE '%" & txtCariDiagnosa.Text & "%') OR (NamaDiagnosa LIKE '%" & txtCariDiagnosa.Text & "%') " _
    & "ORDER BY KdDiagnosa"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgCariDiagnosa.DataSource = rs
    With dgCariDiagnosa
        .Columns(0).Width = 600
        .Columns(1).Width = 10000
    End With

    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadDiagnosa()
    On Error GoTo errLoad
    
    strSQL = "SELECT id, JenisDiagnosa AS [Jenis Diagnosa], Diagnosa FROM RMPDiagnosa " _
    & " WHERE NoPendaftaran='" & mstrNoPen & "' ORDER BY JenisDiagnosa,id"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgDiagnosa.DataSource = rs
    With dgDiagnosa
        .Columns(0).Width = 0
        .Columns(2).Width = 10000
    End With
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadPD()
    On Error GoTo errLoad
    
    strSQL = "SELECT id, Laboratorium, Radiologi, EKG, ProsedurDiagnostik FROM RMPProsedurDiagnostik " _
    & " WHERE NoPendaftaran='" & mstrNoPen & "'"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    If rs.EOF Then Exit Sub
    
    txtLab.Text = Trim(rs("Laboratorium").Value)
    txtRad.Text = Trim(rs("Radiologi").Value)
    txtEKG.Text = Trim(rs("EKG").Value)
    txtPD.Text = Trim(rs("ProsedurDiagnostik").Value)
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadIL()
    On Error GoTo errLoad
    
    strSQL = "SELECT id, KeadaanPulang, AlasanRawat, TujuanKontrol, TanggalKontrol, Rencana FROM RMPInstruksiLanjutan " _
    & " WHERE NoPendaftaran='" & mstrNoPen & "'"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    If rs.EOF Then Exit Sub
    
    txtKeadaanPulang.Text = Trim(rs("KeadaanPulang").Value)
    txtAlasanRawat.Text = Trim(rs("AlasanRawat").Value)
    dcKontrolKe.Text = Trim(rs("TujuanKontrol").Value)
    dtpTglKontrol.Value = Trim(Format(rs("TanggalKontrol").Value, "dd/mm/yyyy 00:00:00"))
    txtRencana.Text = Trim(rs("Rencana").Value)
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadObat()
    On Error GoTo errLoad
    
    strSQL = "SELECT id, Obat FROM RMPObat " _
    & " WHERE NoPendaftaran='" & mstrNoPen & "'"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgObat.DataSource = rs
    With dgObat
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadCariObat()
    On Error GoTo errLoad
    
    strSQL = "SELECT KdBarang, NamaBarang AS [Nama Obat] FROM MasterBarang WHERE " _
    & "NamaBarang LIKE '%" & txtCariObat.Text & "%' AND StatusEnabled = '1' " _
    & "ORDER BY NamaBarang"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgCariObat.DataSource = rs
    With dgCariObat
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With

    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadObatPulang()
    On Error GoTo errLoad
    
    strSQL = "SELECT id, ObatPulang FROM RMPObatPulang " _
    & " WHERE NoPendaftaran='" & mstrNoPen & "'"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgObatPulang.DataSource = rs
    With dgObatPulang
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With
    
    Exit Sub
errLoad:
    msubPesanError
End Sub

Public Sub subLoadCariObatPulang()
    On Error GoTo errLoad
    
    strSQL = "SELECT KdBarang, NamaBarang AS [Nama Obat] FROM MasterBarang WHERE " _
    & "NamaBarang LIKE '%" & txtCariObatPulang.Text & "%' AND StatusEnabled = '1' " _
    & "ORDER BY NamaBarang"

    'Set rs = Nothing
    msubRecFO rs, strSQL
    
    Set dgCariObatPulang.DataSource = rs
    With dgCariObatPulang
        .Columns(0).Width = 0
        .Columns(1).Width = 10000
    End With

    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub sp_simpanRiwayatPenyakit(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("RiwayatPenyakit", adChar, adParamInput, 255, txtRiwayatPenyakit.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPRiwayatPenyakit"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Riwayat Penyakit", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    
    txtRiwayatPenyakit.Text = ""
End Sub

Private Sub sp_simpanPemeriksaanFisik(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("PemeriksaanFisik", adChar, adParamInput, 255, txtPemeriksaanFisik.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPPemeriksaanFisik"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Pemeriksaan Fisik", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    
    txtPemeriksaanFisik.Text = ""
End Sub

Private Sub sp_hapusRiwayatPenyakit(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, dgRiwayatPenyakit.Columns("id").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("RiwayatPenyakit", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPRiwayatPenyakit"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Riwayat Penyakit", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Sub

Private Sub sp_hapusPemeriksaanFisik(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, dgPemeriksaanFisik.Columns("id").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("PemeriksaanFisik", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPPemeriksaanFisik"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Pemeriksaan Fisik", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Sub

Private Sub sp_SimpanDiagnosa(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("Diagnosa", adChar, adParamInput, 255, txtDiagnosa.Text)
        .Parameters.Append .CreateParameter("JenisDiagnosa", adChar, adParamInput, 255, cmbJenisDiagnosa.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Pemeriksaan Fisik", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtDiagnosa.Text = ""
End Sub

Private Sub sp_SimpanPD(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("Laboratorium", adChar, adParamInput, 500, txtLab.Text)
        .Parameters.Append .CreateParameter("Radiologi", adChar, adParamInput, 500, txtRad.Text)
        .Parameters.Append .CreateParameter("EKG", adChar, adParamInput, 500, txtEKG.Text)
        .Parameters.Append .CreateParameter("PD", adChar, adParamInput, 500, txtPD.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPProsedurDiagnostik"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Prosedur Diagnostik", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtDiagnosa.Text = ""
End Sub

Private Sub sp_SimpanIL(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KeadaanPulang", adChar, adParamInput, 500, txtKeadaanPulang.Text)
        .Parameters.Append .CreateParameter("AlasanRawat", adChar, adParamInput, 500, txtAlasanRawat.Text)
        .Parameters.Append .CreateParameter("TujuanKontrol", adChar, adParamInput, 35, dcKontrolKe.Text)
        .Parameters.Append .CreateParameter("TanggalKontrol", adChar, adParamInput, 10, Format(dtpTglKontrol.Value, "yyyy-mm-dd"))
        .Parameters.Append .CreateParameter("Rencana", adChar, adParamInput, 500, txtRencana.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPInstruksiLanjutan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Instruksi Lanjutan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtDiagnosa.Text = ""
End Sub

Private Sub sp_HapusDiagnosa(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, dgDiagnosa.Columns("id").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("Diagnosa", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("JenisDiagnosa", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPDiagnosa"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Diagnosa", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtDiagnosa.Text = ""
End Sub

Private Sub sp_SimpanObat(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("Obat", adChar, adParamInput, 255, txtObat.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPObat"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Obat", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtObat.Text = ""
End Sub

Private Sub sp_HapusObat(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, dgObat.Columns("id").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("Obat", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPObat"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Obat", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtDiagnosa.Text = ""
End Sub

Private Sub sp_SimpanObatPulang(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("ObatPulang", adChar, adParamInput, 255, txtObatPulang.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPObatPulang"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Obat", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtObatPulang.Text = ""
End Sub

Private Sub sp_HapusObatPulang(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("id", adChar, adParamInput, 10, dgObatPulang.Columns("id").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, "")
        .Parameters.Append .CreateParameter("ObatPulang", adChar, adParamInput, 255, "")
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "D")
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RMPObatPulang"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penghapusan Obat", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    txtObatPulang.Text = ""
End Sub

Private Sub txtCariDiagnosa_KeyUp(KeyCode As Integer, Shift As Integer)
    Call subLoadCariDiagnosa
End Sub

Private Sub txtCariObat_KeyUp(KeyCode As Integer, Shift As Integer)
    Call subLoadCariObat
End Sub

Private Sub txtCariObatPulang_KeyUp(KeyCode As Integer, Shift As Integer)
    Call subLoadCariObatPulang
End Sub
