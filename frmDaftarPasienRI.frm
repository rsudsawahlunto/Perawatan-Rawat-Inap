VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarPasienRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Rawat Inap"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienRI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   16455
   Begin VB.CommandButton cmdCariKamarAktif 
      Caption         =   "&Daftar Pasien RI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      TabIndex        =   56
      Top             =   7440
      Width           =   1300
   End
   Begin VB.Frame fraDokterP 
      Caption         =   "Setting Dokter Pemeriksa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   1320
      TabIndex        =   35
      Top             =   2400
      Visible         =   0   'False
      Width           =   12135
      Begin VB.Frame Frame5 
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
         Height          =   975
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   11655
         Begin VB.Frame Frame6 
            Caption         =   "Umur"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8760
            TabIndex        =   43
            Top             =   240
            Width           =   2775
            Begin VB.TextBox txtHr 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               MaxLength       =   6
               TabIndex        =   22
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtBln 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               MaxLength       =   6
               TabIndex        =   21
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtThn 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               MaxLength       =   6
               TabIndex        =   20
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "hr"
               Height          =   210
               Left            =   2400
               TabIndex        =   46
               Top             =   285
               Width           =   165
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "bln"
               Height          =   210
               Left            =   1560
               TabIndex        =   45
               Top             =   285
               Width           =   240
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "thn"
               Height          =   210
               Left            =   720
               TabIndex        =   44
               Top             =   285
               Width           =   285
            End
         End
         Begin VB.TextBox txtJK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            MaxLength       =   9
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtNoPendaftaran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            MaxLength       =   10
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtNoCM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            MaxLength       =   12
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtNamaPasien 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   18
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblJnsKlm 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kelamin"
            Height          =   210
            Left            =   7200
            TabIndex        =   50
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "No. Pendaftaran"
            Height          =   210
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No. CM"
            Height          =   210
            Left            =   1800
            TabIndex        =   48
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblNamaPasien 
            AutoSize        =   -1  'True
            Caption         =   "Nama Pasien"
            Height          =   210
            Left            =   4080
            TabIndex        =   47
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdSimpanDokter 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   7800
         TabIndex        =   28
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton cmdBatalDokter 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   9720
         TabIndex        =   29
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   11655
         Begin VB.TextBox txtPrevDokter 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   5040
            TabIndex        =   25
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtTglPeriksa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2760
            TabIndex        =   24
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtDokter 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   8040
            TabIndex        =   26
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtPoli 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sebelumnya"
            Height          =   210
            Left            =   5040
            TabIndex        =   41
            Top             =   360
            Width           =   2475
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tanggal Pemeriksaan"
            Height          =   210
            Left            =   2760
            TabIndex        =   40
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Dokter Pemeriksa Sekarang"
            Height          =   210
            Left            =   8040
            TabIndex        =   39
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ruang Pemeriksaan"
            Height          =   210
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraDokter 
         Caption         =   "Data Dokter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   240
         TabIndex        =   36
         Top             =   2520
         Width           =   11655
         Begin MSDataGridLib.DataGrid dgDokter 
            Height          =   1935
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3413
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
            RowHeight       =   16
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame fraCari 
      Height          =   840
      Left            =   0
      TabIndex        =   30
      Top             =   7200
      Width           =   16455
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4440
         TabIndex        =   57
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdRencana 
         Caption         =   "Rencana Pindah Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7080
         TabIndex        =   55
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdPesanDarah 
         Caption         =   "Pesan &Darah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "Pesan &Pelayanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4800
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdPesanMenuDiet 
         Caption         =   "Pesan &Menu Gizi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12360
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1300
      End
      Begin VB.CommandButton cmdAsKep 
         Caption         =   "&Asuhan Keperawatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdBatalDirawat 
         Caption         =   "&Batal Dirawat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdUbahRegistrasi 
         Caption         =   "&Ubah Registrasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9720
         TabIndex        =   8
         Top             =   240
         Width           =   1300
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   3135
      End
      Begin VB.CommandButton cmdTP 
         Caption         =   "Transaksi Pela&yanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   11040
         TabIndex        =   9
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdMasukKamar 
         Appearance      =   0  'Flat
         Caption         =   "&Masuk Kamar"
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
         Height          =   450
         Left            =   12480
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdKeluarKamar 
         Appearance      =   0  'Flat
         Caption         =   "&Keluar Kamar"
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
         Height          =   450
         Left            =   13680
         TabIndex        =   11
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15000
         TabIndex        =   15
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   165
         Width           =   2640
      End
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien Rawat Inap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   31
      Top             =   1680
      Width           =   16455
      Begin MSDataGridLib.DataGrid dgDaftarPasienRI 
         Height          =   4695
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   8281
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
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   8550
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   6690
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraPilih 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   34
      Top             =   960
      Width           =   16455
      Begin VB.OptionButton OptRencanaPasien 
         Caption         =   "Daftar Rencana  Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   54
         Top             =   240
         Width           =   3615
      End
      Begin VB.OptionButton optPasNonAktif 
         Caption         =   "Daftar Pasien Pindahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11400
         TabIndex        =   1
         Top             =   200
         Width           =   3735
      End
      Begin VB.OptionButton optPasAktif 
         Caption         =   "Daftar Pasien Aktif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   51
      Top             =   8040
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4727
            MinWidth        =   4409
            Text            =   "Rincian Biaya Perawatan Sementara (F1)"
            TextSave        =   "Rincian Biaya Perawatan Sementara (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4994
            MinWidth        =   4676
            Text            =   "Rincian Kumulatif Biaya Sementara (Ctrl+F1)"
            TextSave        =   "Rincian Kumulatif Biaya Sementara (Ctrl+F1)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2699
            MinWidth        =   2381
            Text            =   "Dokter Pasien (Ctrl+F2)"
            TextSave        =   "Dokter Pasien (Ctrl+F2)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3140
            MinWidth        =   2822
            Text            =   "Ubah Data Pasien (Ctrl+F3)"
            TextSave        =   "Ubah Data Pasien (Ctrl+F3)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2789
            MinWidth        =   2471
            Text            =   "Ubah SMF (Ctrl+F11)"
            TextSave        =   "Ubah SMF (Ctrl+F11)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   0
            Text            =   "Refresh (F5)"
            TextSave        =   "Refresh (F5)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2788
            MinWidth        =   2470
            Text            =   "Ubah Kamar (Ctrl+F7)"
            TextSave        =   "Ubah Kamar (Ctrl+F7)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3317
            MinWidth        =   2999
            Text            =   "Cetak Daf. Pasien (F9)"
            TextSave        =   "Cetak Daf. Pasien (F9)"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4233
            MinWidth        =   4233
            Text            =   "Detail Rincian Biaya Sementara ( F10 )"
            TextSave        =   "Detail Rincian Biaya Sementara ( F10 )"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   53
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienRI.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14655
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12360
      Picture         =   "frmDaftarPasienRI.frx":2328
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmDaftarPasienRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDokter As String
Dim intJmlDokter As Integer
Dim dTglMasuk As Date
Dim bolTampilGrid As Boolean

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcKelas, rs, "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1' order by DeskKelas")
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdAsKep_Click()
    On Error GoTo errLoad
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    With frmAsuhanKeperawatan
        mstrNoPen = dgDaftarPasienRI.Columns("No. Registrasi")
        mstrNoCM = dgDaftarPasienRI.Columns("No. CM")
        .txtnopendaftaran = mstrNoPen
        .txtnocm = mstrNoCM
        .txtNamaPasien = dgDaftarPasienRI.Columns("Nama Pasien")

        If dgDaftarPasienRI.Columns("JK") = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-laki"
        End If
        .txtThn = dgDaftarPasienRI.Columns("UmurTahun")
        .txtBln = dgDaftarPasienRI.Columns("UmurBulan")
        .txtHari = dgDaftarPasienRI.Columns("UmurHari")
        .Show
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatalDirawat_Click()
    On Error GoTo errLoad
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    If MsgBox("Yakin akan membatalkan perawatan pasien " & dgDaftarPasienRI.Columns("Nama Pasien").Value & "", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarPasienRI.Columns("No. Registrasi").Value)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, dgDaftarPasienRI.Columns("No. CM").Value)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, dgDaftarPasienRI.Columns("KdSubInstalasi").Value)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dgDaftarPasienRI.Columns("TglMasuk").Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglBatal", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputMsg", adChar, adParamOutput, 1, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRIBatalDiRawat"
        .CommandType = adCmdStoredProc

        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pembatalan pasien", vbCritical, "Validasi"
        Else
            If LCase(.Parameters("OutputMsg").Value) = "t" Then
                MsgBox "Pelayanan yang didapat di ruangan ini harus dihapus terlebih dahulu", vbCritical, "Validasi"
            Else
                MsgBox "Pasien " & dgDaftarPasienRI.Columns("Nama Pasien").Value & " Batal dirawat", vbInformation, "Informasi"
                Call Add_HistoryLoginActivity("Add_PasienRIBatalDiRawat")
            End If

        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatalDokter_Click()
    fraDokterP.Visible = False
    fraDokterP.Enabled = False
    fraPilih.Enabled = True
    fraDaftar.Enabled = True
    fraCari.Enabled = True
    Call Form_Load

End Sub

Public Sub cmdCari_Click()
    If optPasAktif.Value = True Then
        Call optPasAktif_GotFocus
    Else
        Call optPasNonAktif_GotFocus
    End If
    
    If OptRencanaPasien.Value = True Then OptRencanaPasien_GotFocus
    
End Sub

Private Sub cmdCariKamarAktif_Click()
frmCariPasien.Show
End Sub

Private Sub cmdKeluarKamar_Click()
    On Error GoTo hell
    If dgDaftarPasienRI.Columns(0).Value = "" Then
        Exit Sub
    End If
    
    
    
    mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
    frmDaftarPasienRI.Enabled = False

    '    Edit By Arikawa 2007-09-04
    If sp_UpdateJmlPelayananKamarBK(dgDaftarPasienRI.Columns("No. Registrasi").Value) = False Then Exit Sub

    Call subLoadFormKeluarKam
    Exit Sub
hell:
    Call subLoadFormKeluarKam
    
End Sub

Private Sub cmdMasukKamar_Click()
    On Error GoTo hell
    If dgDaftarPasienRI.Columns(0).Value = "" Then Exit Sub
    mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
    frmDaftarPasienRI.Enabled = False
    Call subLoadFormMasukKam
    Exit Sub
hell:
End Sub

Private Sub cmdOrder_Click()
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    With frmOrderPelayanan
         mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
        .txtnocm = dgDaftarPasienRI.Columns("No. CM").Value
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns("No. Registrasi")

        .Show
    End With
End Sub

Private Sub cmdPesanDarah_Click()
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    With frmPemesananDarah
        .txtnocm = dgDaftarPasienRI.Columns(1).Value
        .txtnopendaftaran = dgDaftarPasienRI.Columns(0).Value
        .txtNamaPasien = dgDaftarPasienRI.Columns("Nama Pasien").Value

        If dgDaftarPasienRI.Columns("JK").Value = "P" Then
            txtJK.Text = "Perempuan"
        Else
            txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
        .txtSubInstalasi.Text = dgDaftarPasienRI.Columns("Subinstalasi").Value
        mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
        .Show
    End With
End Sub

Private Sub cmdPesanMenuDiet_Click()
    On Error GoTo hell

    Call subloadFormPesan
    Exit Sub
hell:
End Sub

Private Sub subloadFormPesan()
    If optPasAktif.Value = True Then
        With frmPesanDiet
            frmDaftarPasienRI.Enabled = False
            strKet = "1"
            .Show
            .dcJenisMenuDiet.SetFocus
        End With
    ElseIf optPasNonAktif.Value = True Then
        MsgBox "Pasien sudah pulang tidak bisa pesan menu", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
End Sub

Private Sub cmdRencana_Click()
 On Error GoTo hell
    If dgDaftarPasienRI.Columns(0).Value = "" Then Exit Sub
'    mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
        
'    frmDaftarPasienRI.Enabled = False

    strSQL = "SELECT  * FROM V_RencanaPindah WHERE  NoPendaftaran ='" & dgDaftarPasienRI.Columns(0) & "' and NoCM='" & dgDaftarPasienRI.Columns(1) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "Pasien Bersangkutan Sudah Terdaftar di Daftar Rencana  Pasien", vbExclamation, "Validasi"
        Exit Sub
    End If
    Call SubLoadFormRencanaPindahPulang
    Exit Sub
hell:
End Sub

Private Sub cmdSimpanDokter_Click()
    On Error GoTo hell
    If mstrKdDokter = "" Then
        MsgBox "Pilih dulu dokternya", vbExclamation, "Validasi"
        txtDokter.SetFocus
        Exit Sub
    End If
    Call sp_UbahDokter(dbcmd)
    Call cmdBatalDokter_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTP_Click()
    On Error GoTo hell
    If dgDaftarPasienRI.Columns(0).Value = "" Then
        Exit Sub
    End If

    If sp_UpdateJmlPelayananKamarBK(dgDaftarPasienRI.Columns("No. Registrasi").Value) = False Then Exit Sub
    Call subLoadFormTP
    Exit Sub
hell:
    MsgBox "Silahkan hubungi Administrator untuk input sewa kamar", vbCritical, "Perhatian"
    Call subLoadFormTP
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub cmdUbahRegistrasi_Click()
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    frmDaftarPasienRI.Enabled = False
    mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
    mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
    Call subLoadFormUbahRegistrasi
End Sub

Private Sub Command1_Click()
cmdCari_Click
End Sub

Private Sub dcJenisPasien_Change()
    If optPasAktif.Value = True Then Call optPasAktif_GotFocus
    If optPasNonAktif.Value = True Then Call optPasNonAktif_GotFocus
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub dcKelas_Change()
    If optPasAktif.Value = True Then Call optPasAktif_GotFocus
    If optPasNonAktif.Value = True Then Call optPasNonAktif_GotFocus
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'Private Sub dgDaftarPasienRI_ButtonClick(ByVal ColIndex As Integer)
'    On Error Resume Next
'
'    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
'    Set rs = Nothing
'    strQuery = "select IdDokter from RegistrasiRI where NoPendaftaran='" & dgDaftarPasienRI.Columns(0).Value & "'"
'    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
'
'    If rs.RecordCount > 0 Then
'        If optPasNonAktif.Value = True Then Exit Sub
'        If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
'        With fraDokterP
'            .Visible = True
'            .Enabled = True
'
'            .Left = (Me.Width - fraDokterP.Width) / 2
'            .Top = fraDaftar.Top
'            .Height = 5415
'        End With
'
'        fraPilih.Enabled = False
'        fraDaftar.Enabled = False
'        fraCari.Enabled = False
'        txtNoPendaftaran.Text = dgDaftarPasienRI.Columns("No. Registrasi").Value
'        txtNoCM.Text = dgDaftarPasienRI.Columns("No. CM").Value
'        txtNamaPasien.Text = dgDaftarPasienRI.Columns("Nama Pasien").Value
'        If dgDaftarPasienRI.Columns("JK").Value = "P" Then
'            txtJK.Text = "Perempuan"
'        Else
'            txtJK.Text = "Laki-Laki"
'        End If
'        txtPoli.Text = strNNamaRuangan
'        txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
'        txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
'        txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
'        dTglMasuk = dgDaftarPasienRI.Columns("TglMasuk").Value
'        txtTglPeriksa.Text = Format(dTglMasuk, "dd MMMM yyyy HH:mm:ss")
'        txtDokter.Text = ""
'        txtDokter.SetFocus
'        strSQL = "select DokterPenanggungJawab from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and NoPendaftaran='" & txtNoPendaftaran.Text & "'"
'        Set rs = Nothing
'        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'        If IsNull(rs("DokterPenanggungJawab").Value) Then
'            txtPrevDokter.Text = ""
'        Else
'            txtPrevDokter.Text = rs("DokterPenanggungJawab").Value
'        End If
'    End If
'End Sub

Private Sub dgDaftarPasienRI_Click()
    bolTampilGrid = True
    WheelHook.WheelUnHook
    Set MyProperty = dgDaftarPasienRI
    WheelHook.WheelHook dgDaftarPasienRI
End Sub

Private Sub dgDaftarPasienRI_DblClick()

    If mstrKdRuangan = "325" Then
        With frmDataPasienBayiTabung
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = dgDaftarPasienRI.Columns(1).Value
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "L" Then
                .txtJK.Text = "Laki-Laki"
            Else
                .txtJK.Text = "Perempuan"
            End If

            .txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
            .txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
            .txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
            .txtSubInstalasi.Text = dgDaftarPasienRI.Columns("Subinstalasi").Value
            .dcSubInstalasi.Text = dgDaftarPasienRI.Columns("Subinstalasi").Value

            mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value

        End With
    End If
End Sub

Private Sub dgDaftarPasienRI_GotFocus()
    dgDaftarPasienRI.Refresh
    dgDaftarPasienRI.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub dgDaftarPasienRI_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 0
            mstrFilter = " Order By NoPendaftaran"
        Case 1
            mstrFilter = " Order By NoCM"
        Case 2
            mstrFilter = " Order By [Nama Pasien]"
        Case 3
            mstrFilter = " Order By JK"
        Case 4
            mstrFilter = " Order By Umur"
        Case 5
            mstrFilter = " Order By Kelas"
        Case 6
            mstrFilter = " Order By JenisPasien"
        Case Else
            mstrFilter = ""
    End Select
    If optPasAktif.Value = True Then Call optPasAktif_GotFocus
    If optPasNonAktif.Value = True Then Call optPasNonAktif_GotFocus
End Sub

Private Sub dgDaftarPasienRI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optPasAktif.Value = True Then cmdTP.SetFocus Else cmdMasukKamar.SetFocus
    End If
End Sub

Private Sub dgDaftarPasienRI_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next

    If bolTampilGrid = False Then Exit Sub
    If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
    Set rs = Nothing
'    strQuery = "select IdDokter from RegistrasiRI where NoPendaftaran='" & dgDaftarPasienRI.Columns(0).Value & "'"
'    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
'    If IsNull(rs.Fields(0).Value) = True Then Exit Sub

    lblJumData.Caption = "Data " & dgDaftarPasienRI.Bookmark & "/" & dgDaftarPasienRI.ApproxCount
    dgDaftarPasienRI.Refresh

End Sub

Private Sub dgDokter_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDokter
    WheelHook.WheelHook dgDokter
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        cmdSimpanDokter.SetFocus
    End If
End Sub

Public Sub PostingHutangPenjaminPasien_AU(strStatus As String)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgDaftarPasienRI.Columns("No. Registrasi").Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.PostingHutangPenjaminPasien_AU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam proses update HP pasien", vbCritical, "Validasi"
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing

    Exit Sub
hell_:
    msubPesanError
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errLoad
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode

        Case vbKeyF1
            If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
            mstrTglKeluar = Now
            mstrNoPen = dgDaftarPasienRI.Columns(0).Value
            If sp_UpdateJmlPelayananKamarBK(mstrNoPen) = False Then Exit Sub
            strCetak = IIf(strCtrlKey = "2", "Lengkap", "Singkat") ' Cek Ctrl+F1
            mstrFormPengirim = Me.Name
            frmDaftarPasienRI.Enabled = False
            frmFilterReport.Show
            '  frm_cetak_RincianBiaya.Show

        Case vbKeyF2
            If strCtrlKey = 4 Then
                If OptRencanaPasien.Value = True Then Exit Sub
                If optPasNonAktif.Value = True Then Exit Sub
                If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
                With fraDokterP
                    .Visible = True
                    .Enabled = True

                    .Left = (Me.Width - fraDokterP.Width) / 2
                    .Top = fraDaftar.Top
                    .Height = 5415
                End With

                fraPilih.Enabled = False
                fraDaftar.Enabled = False
                fraCari.Enabled = False
                txtnopendaftaran.Text = dgDaftarPasienRI.Columns("No. Registrasi").Value
                txtnocm.Text = dgDaftarPasienRI.Columns("No. CM").Value
                txtNamaPasien.Text = dgDaftarPasienRI.Columns("Nama Pasien").Value
                If dgDaftarPasienRI.Columns("JK").Value = "P" Then
                    txtJK.Text = "Perempuan"
                Else
                    txtJK.Text = "Laki-Laki"
                End If
                txtPoli.Text = strNNamaRuangan
                txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
                txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
                txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
                dTglMasuk = dgDaftarPasienRI.Columns("TglMasuk").Value
                txtTglPeriksa.Text = Format(dTglMasuk, "dd MMMM yyyy HH:mm:ss")
                txtDokter.Text = ""
                txtDokter.SetFocus
                strSQL = "select DokterPenanggungJawab from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and NoPendaftaran='" & txtnopendaftaran.Text & "'"
                Set rs = Nothing
                rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                If IsNull(rs("DokterPenanggungJawab").Value) Then
                    txtPrevDokter.Text = ""
                Else
                    txtPrevDokter.Text = rs("DokterPenanggungJawab").Value
                End If
            End If
            Set frmDaftarPasienRI = Nothing
        Case vbKeyF3
            If strCtrlKey = 4 Then
                If OptRencanaPasien.Value = True Then Exit Sub
                If optPasNonAktif.Value = True Then Exit Sub
                strPasien = "Lama"
                If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
                mstrNoCM = dgDaftarPasienRI.Columns(1).Value

                frmPasienBaru.Show
            End If

        Case vbKeyF9
            frmCtkDaftarPasien.Show

        Case vbKeyF10
            If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
            mstrNoPen = dgDaftarPasienRI.Columns("No. Registrasi").Value
'            If dgDaftarPasienRI.Columns("JenisPasien") <> "UMUM" Then Call PostingHutangPenjaminPasien_AU("A")
'            frm_cetak_RincianBiayaPenjamin.Show
            
            If frmDaftarPasienRI.OptRencanaPasien.Value = True Then
                frm_cetak_RincianBiayaPenjamin.Show
            ElseIf frmDaftarPasienRI.optPasAktif.Value = True Then
                If frmDaftarPasienRI.dgDaftarPasienRI.Columns("JenisPasien") <> "UMUM" Then Call frmDaftarPasienRI.PostingHutangPenjaminPasien_AU("A")
                frm_cetak_RincianBiayaPenjamin.Show
            End If

        Case vbKeyF11
            If strCtrlKey = 4 Then
                If OptRencanaPasien.Value = True Then Exit Sub
                If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
                frmDaftarPasienRI.Enabled = False
                mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
                mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
                mstrNoPen = dgDaftarPasienRI.Columns(0).Value
                mstrNoCM = dgDaftarPasienRI.Columns(1).Value
                mdTglMasuk = dgDaftarPasienRI.Columns("tglMasuk").Value

                With frmUbahKasusPenyakitPasien
                    .Show
                    .txtnopendaftaran.Text = mstrNoPen
                    .txtNamaFormPengirim.Text = Me.Name
                    .txtNoPendaftaran_KeyPress (13)
                End With
            End If

        Case vbKeyF5
            Call cmdCari_Click

        Case vbKeyF7
            If strCtrlKey = 4 Then
                If OptRencanaPasien.Value = True Then Exit Sub
                If dgDaftarPasienRI.ApproxCount = 0 Then Exit Sub
                frmDaftarPasienRI.Enabled = False
                mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
                mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
                Call subLoadFormUbahKamarBed
            End If

    End Select

    Exit Sub
errLoad:
    frmDaftarPasienRI.Enabled = True
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    Call subLoadDcSource
    mstrFilter = ""
    optPasAktif.Caption = "Daftar Pasien " & strNNamaRuangan
    optPasAktif.Value = True
    Set rs = Nothing
    Call cmdCari_Click
    mblnForm = True
    mblnFormDaftarPasienRI = True
    bolTampilGrid = False
    If mblnAdmin = False Then
        cmdUbahRegistrasi.Enabled = False
        cmdBatalDirawat.Enabled = False
    Else
        cmdUbahRegistrasi.Enabled = True
        cmdBatalDirawat.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnForm = False
    mblnFormDaftarPasienRI = False
End Sub

Public Sub optPasAktif_GotFocus()
    On Error GoTo hell

    lblJumData.Caption = "Data 0/0"
    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,SubInstalasi,KdSubInstalasi,KdKelas,KdRuangan,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%')  AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    lblJumData.Caption = "Data 0/" & rs.RecordCount
    Set dgDaftarPasienRI.DataSource = rs
    Call SetGridPasienRIAktif
    cmdKeluarKamar.Enabled = True
    cmdMasukKamar.Enabled = False
    cmdPesanMenuDiet.Visible = True
    cmdPesanMenuDiet.Enabled = True
    cmdTP.Enabled = True
    cmdAsKep.Enabled = True
    cmdPesanDarah.Enabled = True
    cmdOrder.Enabled = True
    cmdRencana.Enabled = True
    
    
    
    StatusBar1.Panels(1).Visible = True
    StatusBar1.Panels(2).Visible = True
    StatusBar1.Panels(3).Visible = True
    StatusBar1.Panels(4).Visible = True
    StatusBar1.Panels(5).Visible = True
    'StatusBar1.Panels(6).Visible = False
    StatusBar1.Panels(7).Visible = True
    StatusBar1.Panels(8).Visible = True

    If mblnAdmin = False Then
        cmdUbahRegistrasi.Enabled = False
        cmdBatalDirawat.Enabled = False
    Else
        cmdUbahRegistrasi.Enabled = True
        cmdBatalDirawat.Enabled = True
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub optPasAktif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgDaftarPasienRI.SetFocus
End Sub

Public Sub optPasNonAktif_GotFocus()
    On Error GoTo hell
    lblJumData.Caption = "Data 0/0"
    Set rs = Nothing
    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,[Ruangan Asal],[Ruangan Tujuan],TglPindah,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,KdRuanganTujuan,KdRuangan from V_DaftarPasienRIPindahKamar where ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%') AND KdRuanganTujuan='" & mstrKdRuangan & "' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%'" & mstrFilter
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRI.DataSource = rs
    Call SetGridPasienRINonAktif
    cmdKeluarKamar.Enabled = False
    cmdMasukKamar.Enabled = True
    cmdUbahRegistrasi.Enabled = False
    cmdPesanMenuDiet.Enabled = False
    cmdPesanMenuDiet.Visible = False
    cmdTP.Enabled = False
    cmdAsKep.Enabled = False
    cmdBatalDirawat.Enabled = False
    cmdPesanDarah.Enabled = False
    cmdOrder.Enabled = False
    cmdRencana.Enabled = False
    
    StatusBar1.Panels(1).Visible = False
    StatusBar1.Panels(2).Visible = False
    StatusBar1.Panels(3).Visible = False
    StatusBar1.Panels(4).Visible = False
    StatusBar1.Panels(5).Visible = False
    'StatusBar1.Panels(6).Visible = False
    StatusBar1.Panels(7).Visible = False
    StatusBar1.Panels(8).Visible = False
    StatusBar1.Panels(9).Visible = False
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub optPasNonAktif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dgDaftarPasienRI.SetFocus
End Sub

Private Sub SetGridRencanaPasienPulang()
With dgDaftarPasienRI
        '        .AllowRowSizing = False 'new
        
        
       '.Columns(0).Button = True
        .Columns(0).Caption = "No. Registrasi"
        .Columns(0).Width = 1500 '1250
        
        
        .Columns(1).Caption = "No. CM"
        .Columns(1).Alignment = dbgCenter
        .Columns(1).Width = 1000
        .Columns(2).Width = 1800
        .Columns(3).Width = 300
        .Columns(4).Caption = "Umur"
        .Columns(4).Width = 600
        .Columns(5).Width = 2000
        .Columns(6).Width = 2000
        .Columns(15).Width = 0 'kdStatusKeluar
        
        
'        If .Columns(15).Value = "01" Then
            .Columns(9).Width = 1200
            .Columns(8).Width = 1900 ' namaTempatTujuan
            .Columns(22).Width = 1300 ' Kelas Pelayanan
'        Else
'            .Columns(9).Width = 1200 ' Status Pulang
'            .Columns(8).Width = 0 ' namaTempatTujuan
'            .Columns(22).Width = 0 ' Kelas Pelayanan
'        End If
        .Columns(7).Width = 1700 ' status keluar
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
        .Columns(19).Width = 0
        .Columns(20).Width = 0
        .Columns(21).Width = 0
        
 End With

End Sub


Public Sub OptRencanaPasien_GotFocus()
 On Error GoTo hell

    lblJumData.Caption = "Data 0/0"
    Set rs = Nothing
'    strQuery = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,SubInstalasi,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%')  AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND Kelas LIKE '%" & dcKelas.Text & "%'" & mstrFilter
    strQuery = "SELECT  NoPendaftaran,NoCM,[Nama Pasien],JK,UmurTahun,TglMasuk,tglRencanaKeluar,StatusKeluar,NamaTempatTujuan,StatusPulang,Nopakai,kdKondisiPulang,KdRuanganTujuan,KdRuanganAsal,KdKelas,KdStatusKeluar,KdStatusPulang,KdSubInstalasi,NoOrder,UmurTahun,UmurBulan,UmurHari,Kelas " & _
               " FROM V_RencanaPindah WHERE kdruangan='" & mstrKdRuangan & "' and ([Nama Pasien] like '%" & txtParameter.Text & "%' or NoCM like '%" & txtParameter.Text & "%') "
    rs.Open strQuery, dbConn, adOpenStatic, adLockOptimistic
    Set dgDaftarPasienRI.DataSource = rs
    If rs.RecordCount > 0 Then rs.MoveFirst
   
    Call SetGridRencanaPasienPulang
    cmdKeluarKamar.Enabled = True
    cmdMasukKamar.Enabled = False
    cmdPesanMenuDiet.Enabled = False
    cmdTP.Enabled = False
    cmdAsKep.Enabled = False
    cmdBatalDirawat.Enabled = False
    cmdUbahRegistrasi.Enabled = False
    cmdPesanDarah.Enabled = False
    cmdOrder.Enabled = False
    cmdRencana.Enabled = False
    
    
    
    StatusBar1.Panels(1).Visible = True
    StatusBar1.Panels(2).Visible = True
    StatusBar1.Panels(3).Visible = False
    StatusBar1.Panels(4).Visible = False
    StatusBar1.Panels(5).Visible = False
    'StatusBar1.Panels(6).Visible = False
    StatusBar1.Panels(7).Visible = False
    StatusBar1.Panels(8).Visible = True
    StatusBar1.Panels(9).Visible = True
'    If mblnAdmin = False Then
'        cmdUbahRegistrasi.Enabled = False
'        cmdBatalDirawat.Enabled = False
'    Else
'        cmdUbahRegistrasi.Enabled = True
'        cmdBatalDirawat.Enabled = True
'    End If

    Exit Sub
hell:
    Call msubPesanError
 
End Sub

Private Sub txtDokter_Change()
    strFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_GotFocus()
    If txtDokter.Text = "" Then strFilterDokter = ""
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        dgDokter.SetFocus
    End If
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'untuk set grid pasien ri aktif
Sub SetGridPasienRIAktif()
    With dgDaftarPasienRI
        '        .AllowRowSizing = False 'new
        '.Columns(0).Button = True
        .Columns(0).Width = 1300 '1250
        .Columns(0).Alignment = dbgCenter
        .Columns(0).Caption = "No. Registrasi"
        .Columns(1).Width = 1750
        .Columns(1).Caption = "No. CM"
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2000
        .Columns(3).Width = 400
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 1500
        .Columns(5).Width = 1600
        .Columns(6).Width = 1600
        .Columns(7).Width = 1590
        .Columns(8).Width = 1300
        .Columns(8).Caption = "  No. Kamar"
        .Columns(8).Alignment = dbgCenter
        .Columns(9).Width = 580
        .Columns(9).Alignment = dbgCenter
        .Columns(10).Width = 0
        .Columns(11).Width = 0
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 1900
        
    End With
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
        '.Columns(7).Caption = "Tgl. Pindah"
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

'Untuk load data pasien di form Pesan Menu Diet
Private Sub subloadFormPesanMenu()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    If optPasAktif.Value = True Then

        With frmPesanMenuDiet2
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = dgDaftarPasienRI.Columns(1).Value
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
                .txtSex.Text = "Perempuan"
            Else
                .txtSex.Text = "Laki-Laki"
            End If
            .txtKls.Text = dgDaftarPasienRI.Columns("Kelas").Value
            .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(13).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(7).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(7).Value
            mstrKdKelas = dgDaftarPasienRI.Columns(15).Value
            strNoPakai = dgDaftarPasienRI.Columns(10).Value
            mstrKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns(14)
        End With
    ElseIf optPasNonAktif.Value = True Then
        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
            Me.Enabled = True
            Exit Sub
        End If

        With frmPesanMenuDiet
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = mstrNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            .txtSex.Text = dgDaftarPasienRI.Columns(3).Value
            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(11).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(12).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(12).Value
            mstrKdKelas = dgDaftarPasienRI.Columns(14).Value
            mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
        End With
    End If

    Exit Sub
hell:
End Sub

'untuk load data pasien di form transaksi pasien
Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value

    If optPasAktif.Value = True Then
        With frmTransaksiPasien
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = dgDaftarPasienRI.Columns(1).Value
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
                .txtSex.Text = "Perempuan"
            Else
                .txtSex.Text = "Laki-Laki"
            End If
            .txtKls.Text = dgDaftarPasienRI.Columns("Kelas").Value
            .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(13).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(7).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(7).Value
            mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
            mstrKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns("KdSubInstalasi")
        End With
    ElseIf optPasNonAktif.Value = True Then
        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
            Me.Enabled = True
            Exit Sub
        End If
        With frmTransaksiPasien
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = mstrNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            .txtSex.Text = dgDaftarPasienRI.Columns(3).Value
            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
            .txtHr.Text = dgDaftarPasienRI.Columns(11).Value
            .txtJenisPasien.Text = dgDaftarPasienRI.Columns(6).Value
            .txtTglDaftar.Text = dgDaftarPasienRI.Columns(12).Value
            mdTglMasuk = dgDaftarPasienRI.Columns(12).Value
            mstrKdKelas = dgDaftarPasienRI.Columns("KdKelas").Value
            mstrKdSubInstalasi = dgDaftarPasienRI.Columns("KdSubInstalasi").Value
        End With
    End If

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    Exit Sub
hell:
End Sub

'untuk load data pasien di form keluar kamar
Private Sub subLoadFormKeluarKam()
    On Error GoTo hell
'    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
'    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    With frmKeluarKamar
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtnocm.Text = mstrNoCM
         .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
        .txtHari.Text = dgDaftarPasienRI.Columns("UmurHari").Value
'        .txtNoPemakaian.Text = dgDaftarPasienRI.Columns(10).Value
          .txtNoPemakaian.Text = dgDaftarPasienRI.Columns("noPakai").Value
        .txtTglMasuk.Text = dgDaftarPasienRI.Columns("TglMasuk").Value
'        .txtKeterangan.Text = dgDaftarPasienRI.Columns(17).Value
    End With
    Exit Sub
hell:
End Sub
 
 Private Sub SubLoadFormRencanaPindahPulang()
  On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    mstrNoPakai = dgDaftarPasienRI.Columns("noPakai").Value
'    mstrKdKelas = dgDaftarPasienRI.Columns("KdKelas").Value
'    strSQL = "SELECT * FROM PemakaianKamar WHERE NoCM='" & mstrNoCM _
'    & "' AND StatusKeluar='T'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    If rs.RecordCount <> 0 Then
'        MsgBox "Pasien belum keluar kamar", vbCritical, "Validasi"
'        Me.Enabled = True
'        Exit Sub
'    End If
    With frmRencanaPindahPulangPasien
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienRI.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienRI.Columns("UmurHari").Value
        .dcKelas.Text = dgDaftarPasienRI.Columns("kelas").Value
        .txtKdRuanganAsal.Text = dgDaftarPasienRI.Columns("KdRuangan").Value
        
        .dtpTglRencanaKeluar.SetFocus
    End With
    Exit Sub
hell:
Call msubPesanError
'Resume 0
End Sub



'untuk load data pasien di form masuk kamar
Private Sub subLoadFormMasukKam()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    mstrKdKelas = dgDaftarPasienRI.Columns("KdKelas").Value
    strSQL = "SELECT * FROM PemakaianKamar WHERE NoCM='" & mstrNoCM _
    & "' AND StatusKeluar='T'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar kamar", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    With frmMasukKamar
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(10).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(11).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(12).Value

        .txtKdRuanganAsal.Text = dgDaftarPasienRI.Columns(16).Value
    End With
    Exit Sub
hell:
End Sub

'untuk load data pasien di form masuk kamar
Private Sub subLoadFormPsnPulang()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    strSQL = "SELECT * FROM PemakaianKamar WHERE NoCM='" & mstrNoCM _
    & "' AND StatusKeluar='T'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar kamar", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
        MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
        Me.Enabled = True
        Exit Sub
    End If
    With frmPasienPulang
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(11).Value
        .txtTglMasuk.Text = dgDaftarPasienRI.Columns(7).Value
    End With
    Exit Sub
hell:
End Sub

'untuk load data pasien di form transaksi pasien
Private Sub subLoadFormPeriksaDiagnosa()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    If optPasAktif.Value = True Then
        With frmPeriksaDiagnosa
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = mstrNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
                .txtSex.Text = "Perempuan"
            Else
                .txtSex.Text = "Laki-Laki"
            End If
            .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
            .txtHari.Text = dgDaftarPasienRI.Columns(13).Value
            strKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns(14)
        End With
    ElseIf optPasNonAktif.Value = True Then
        If dgDaftarPasienRI.Columns(8).Value <> mstrNamaRuangan Then
            MsgBox "Anda tidak berhak mengakses pasien dari ruangan lain", vbCritical, "Validasi"
            Me.Enabled = True
            Exit Sub
        End If
        With frmPeriksaDiagnosa
            .Show
            .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
            .txtnocm.Text = mstrNoCM
            .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
            If dgDaftarPasienRI.Columns(3).Value = "P" Then
                .txtSex.Text = "Perempuan"
            Else
                .txtSex.Text = "Laki-Laki"
            End If
            .txtThn.Text = dgDaftarPasienRI.Columns(9).Value
            .txtBln.Text = dgDaftarPasienRI.Columns(10).Value
            .txtHari.Text = dgDaftarPasienRI.Columns(11).Value
            strKdSubInstalasi = frmDaftarPasienRI.dgDaftarPasienRI.Columns(13)
        End With
    End If
    Exit Sub
hell:
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & strFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 1200
        .Columns(1).Width = 3500
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_UbahDokter(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dTglMasuk, "yyyy/MM/dd HH:mm:ss"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_DokterPemeriksaRI"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Dokter Pemeriksa pasien", vbCritical, "Validasi"
        Else
            'MsgBox "Penyimpanan Dokter Pemeriksa pasien sukses", vbInformation, "Informasi"
            Call Add_HistoryLoginActivity("Update_DokterPemeriksaRI")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

'untuk load data pasien di form ubah registrasi pasien
Private Sub subLoadFormUbahRegistrasi()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value

    With frmUbahRegistrasiPasienRI
        .txtnopendaftaran.Text = mstrNoPen
        .dcKelasKamar.BoundText = mstrKdKelas
        .txtIdDokterBaru.Text = ""
        .txtNoPendaftaran_KeyPress (13)
        .Show
    End With
    Exit Sub
hell:
End Sub

'untuk load data pasien di form ubah kamar dan bed pasien
Private Sub subLoadFormUbahKamarBed()
    On Error GoTo hell

    mstrNoPen = dgDaftarPasienRI.Columns(0).Value
    mstrNoCM = dgDaftarPasienRI.Columns(1).Value
    mstrKdKelas = dgDaftarPasienRI.Columns(16).Value
    With frmUbahKamardanBed
        .Show
        .txtNoPakai.Text = dgDaftarPasienRI.Columns("NoPakai").Value
        .txtnopendaftaran.Text = dgDaftarPasienRI.Columns(0).Value
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasienRI.Columns(2).Value
        If dgDaftarPasienRI.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgDaftarPasienRI.Columns(11).Value
        .txtBln.Text = dgDaftarPasienRI.Columns(12).Value
        .txtHari.Text = dgDaftarPasienRI.Columns(13).Value

        .txtKdRuanganAsal.Text = dgDaftarPasienRI.Columns(17).Value
        .txtRuangPerawatan.Text = mstrNamaRuangan
        .dcKelasPK.BoundText = mstrKdKelas
        .txtNoKamLama.Text = dgDaftarPasienRI.Columns(8)
        .txtNoBedLama.Text = dgDaftarPasienRI.Columns("NoBed")
    End With
    Exit Sub
hell:
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call cmdCari_Click
        txtParameter.SetFocus
    End If
End Sub

Private Function sp_UpdateJmlPelayananKamarBK(f_NoPendaftaran As String) As Boolean
    sp_UpdateJmlPelayananKamarBK = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_JmlPelayananKamarBKNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan saat update Jumlah Pelayanan", vbCritical, "Validasi"
            sp_UpdateJmlPelayananKamarBK = False
        Else
            Call Add_HistoryLoginActivity("Update_JmlPelayananKamarBKNew")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

