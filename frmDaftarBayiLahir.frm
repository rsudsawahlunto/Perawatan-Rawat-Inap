VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDaftarBayiLahir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Bersalin/Bayi Lahir"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarBayiLahir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14760
   Begin VB.CommandButton cmdEventBayiLahir 
      Caption         =   "E&vent Bayi Lahir"
      Height          =   570
      Left            =   9600
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien Bersalin/Bayi Lahir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   35
      Top             =   1560
      Width           =   14775
      Begin VB.Frame fraEventBayi 
         Caption         =   "Data Event Bayi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   840
         TabIndex        =   55
         Top             =   5400
         Visible         =   0   'False
         Width           =   9375
         Begin VB.CommandButton cmdSimpanEventBayi 
            Caption         =   "&Simpan"
            Height          =   570
            Left            =   4800
            TabIndex        =   12
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton cmdTutupEventBayi 
            Caption         =   "Tutu&p"
            Height          =   570
            Left            =   6300
            TabIndex        =   13
            Top             =   1800
            Width           =   1455
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
            Height          =   1215
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   9015
            Begin MSDataListLib.DataCombo dcNamaEvent 
               Height          =   330
               Left            =   2040
               TabIndex        =   10
               Top             =   480
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.TextBox txtNourutBayi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   9
               Top             =   480
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtpTglEventBayi 
               Height          =   375
               Left            =   5400
               TabIndex        =   11
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy HH:mm"
               Format          =   131989507
               UpDown          =   -1  'True
               CurrentDate     =   38212
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "No. Urut Bayi"
               Height          =   210
               Left            =   240
               TabIndex        =   63
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Tanggal Event Bayi"
               Height          =   210
               Left            =   5400
               TabIndex        =   62
               Top             =   240
               Width           =   1560
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Nama Event Bayi"
               Height          =   210
               Left            =   2040
               TabIndex        =   61
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   8400
            TabIndex        =   59
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Tutup"
            Height          =   375
            Left            =   10440
            TabIndex        =   58
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdSimpanBayiLahir 
            Caption         =   "&Simpan"
            Height          =   570
            Left            =   8400
            TabIndex        =   57
            Top             =   3480
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Tutup"
            Height          =   570
            Left            =   9900
            TabIndex        =   56
            Top             =   3480
            Width           =   1455
         End
      End
      Begin VB.Frame fraBayiLahir 
         Caption         =   "Data Bayi Lahir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   960
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   13215
         Begin VB.CommandButton cmdTutupBayiLahir 
            Caption         =   "Tutu&p"
            Height          =   570
            Left            =   11520
            TabIndex        =   32
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "&Simpan"
            Height          =   570
            Left            =   9960
            TabIndex        =   31
            Top             =   3600
            Width           =   1455
         End
         Begin VB.CommandButton cmdBatalDokter 
            Caption         =   "&Tutup"
            Height          =   375
            Left            =   10440
            TabIndex        =   48
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdSimpanDokter 
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   8400
            TabIndex        =   47
            Top             =   5160
            Width           =   1935
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   12855
            Begin VB.TextBox txtKeterangan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4440
               MaxLength       =   9
               TabIndex        =   30
               Top             =   2400
               Width           =   6375
            End
            Begin VB.TextBox txtNoCMBayi 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9240
               MaxLength       =   9
               TabIndex        =   28
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtNamaLengkap 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   240
               MaxLength       =   9
               TabIndex        =   19
               Top             =   1200
               Width           =   3375
            End
            Begin VB.ComboBox cbJK 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   3720
               TabIndex        =   16
               Text            =   "Laki-Laki"
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox txtKelainan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               MaxLength       =   9
               TabIndex        =   29
               Top             =   2400
               Width           =   4095
            End
            Begin VB.TextBox txtWarnaKulit 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3720
               MaxLength       =   9
               TabIndex        =   20
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox txtTinggiBadan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9120
               MaxLength       =   9
               TabIndex        =   18
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtNoCM 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   15
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtNoPendaftaran 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               MaxLength       =   10
               TabIndex        =   14
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txtBeratBadan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6240
               MaxLength       =   9
               TabIndex        =   17
               Top             =   480
               Width           =   2415
            End
            Begin MSComCtl2.DTPicker dtpTglLahirBayi 
               Height          =   375
               Left            =   6840
               TabIndex        =   27
               Top             =   1800
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd MMM yyyy HH:mm"
               Format          =   116588547
               UpDown          =   -1  'True
               CurrentDate     =   38212
            End
            Begin MSDataListLib.DataCombo dcKondisiBayi 
               Height          =   330
               Left            =   6240
               TabIndex        =   21
               Top             =   1200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcKuantitas 
               Height          =   330
               Left            =   10560
               TabIndex        =   23
               Top             =   1200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcLetakJanin 
               Height          =   330
               Left            =   240
               TabIndex        =   24
               Top             =   1800
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcCaraLahir 
               Height          =   330
               Left            =   2400
               TabIndex        =   25
               Top             =   1800
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcTempatLahirBayi 
               Height          =   330
               Left            =   4560
               TabIndex        =   26
               Top             =   1800
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dcPenyebabKematian 
               Height          =   330
               Left            =   8400
               TabIndex        =   22
               Top             =   1200
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               Text            =   ""
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Penyebab Kematian"
               Height          =   210
               Left            =   8400
               TabIndex        =   72
               Top             =   960
               Width           =   1620
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Keterangan"
               Height          =   210
               Left            =   4440
               TabIndex        =   71
               Top             =   2160
               Width           =   945
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "No.CM Bayi"
               Height          =   210
               Left            =   9240
               TabIndex        =   70
               Top             =   1560
               Width           =   900
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Tempat Lahir Bayi"
               Height          =   210
               Left            =   4560
               TabIndex        =   69
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Cara Lahir Bayi"
               Height          =   210
               Left            =   2400
               TabIndex        =   68
               Top             =   1560
               Width           =   1155
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Letak Janin Bayi"
               Height          =   210
               Left            =   240
               TabIndex        =   67
               Top             =   1560
               Width           =   1290
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Kuantitas Bayi"
               Height          =   210
               Left            =   10560
               TabIndex        =   66
               Top             =   960
               Width           =   1125
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Nama Lengkap"
               Height          =   210
               Left            =   240
               TabIndex        =   65
               Top             =   960
               Width           =   1200
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Kelamin"
               Height          =   210
               Left            =   3720
               TabIndex        =   54
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Tanggal Lahir"
               Height          =   210
               Left            =   6840
               TabIndex        =   53
               Top             =   1560
               Width           =   1080
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Kelainan"
               Height          =   210
               Left            =   240
               TabIndex        =   52
               Top             =   2160
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kondisi Lahir"
               Height          =   210
               Left            =   6240
               TabIndex        =   51
               Top             =   960
               Width           =   990
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Warna Kulit"
               Height          =   210
               Left            =   3720
               TabIndex        =   50
               Top             =   960
               Width           =   930
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tinggi Badan(cm)"
               Height          =   210
               Left            =   9120
               TabIndex        =   49
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "No. CM"
               Height          =   210
               Left            =   1920
               TabIndex        =   46
               Top             =   240
               Width           =   585
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "No. Pendaftaran"
               Height          =   210
               Left            =   240
               TabIndex        =   45
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblJnsKlm 
               AutoSize        =   -1  'True
               Caption         =   "Berat Badan(gr)"
               Height          =   210
               Left            =   6240
               TabIndex        =   44
               Top             =   240
               Width           =   1305
            End
         End
      End
      Begin VB.Frame fraPeriode 
         Caption         =   "Periode"
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
         Left            =   8880
         TabIndex        =   38
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   116391939
            UpDown          =   -1  'True
            CurrentDate     =   38212
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   116391939
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   39
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgDaftarPasienGD 
         Height          =   4815
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   8493
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
      Begin MSDataListLib.DataCombo dcJenisPasien 
         Height          =   330
         Left            =   5040
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtTempNourutBayi 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   8385
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   4057
            MinWidth        =   1411
            Text            =   "Rincian Biaya Pelayanan (F1)"
            TextSave        =   "Rincian Biaya Pelayanan (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   7030
            MinWidth        =   3352
            Text            =   "Rincian Biaya Pelayanan Kumulatif (Ctrl+F1)"
            TextSave        =   "Rincian Biaya Pelayanan Kumulatif (Ctrl+F1)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   6563
            MinWidth        =   530
            Text            =   "Cetak Label (Shift+F1)"
            TextSave        =   "Cetak Label (Shift+F1)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   4763
            MinWidth        =   4763
            Text            =   "Pasien Selesai di Periksa (Ctrl+D)"
            TextSave        =   "Pasien Selesai di Periksa (Ctrl+D)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   13485
            MinWidth        =   1764
            Text            =   "Cetak Surat Keterangan (Ctrl+Z)"
            TextSave        =   "Cetak Surat Keterangan (Ctrl+Z)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   25973
            MinWidth        =   705
            Text            =   "Cetak Daftar Pasien (F9)"
            TextSave        =   "Cetak Daftar Pasien (F9)"
         EndProperty
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
   End
   Begin VB.Frame fraPilih 
      Height          =   615
      Left            =   0
      TabIndex        =   36
      Top             =   960
      Width           =   14775
   End
   Begin VB.Frame fraCari 
      Caption         =   "Cari Data Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   33
      Top             =   7440
      Width           =   14775
      Begin VB.CommandButton cmdBayiLahir 
         Caption         =   "Ba&yi Lahir"
         Height          =   570
         Left            =   11700
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   570
         Left            =   13200
         TabIndex        =   8
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   450
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien /  No.CM"
         Height          =   210
         Left            =   840
         TabIndex        =   34
         Top             =   195
         Width           =   2640
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   41
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
      Left            =   12960
      Picture         =   "frmDaftarBayiLahir.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarBayiLahir.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarBayiLahir.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmDaftarBayiLahir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim strFilterDokter As String
Dim dTglMasuk As Date
Dim strNoHasilPeriksa As String

'Store procedure untuk mengisi registrasi pasien
Private Function sp_AUD_DetailHasilTindakanMedisBersalin(f_status As String) As Boolean
Set dbcmd = New ADODB.Command
sp_AddBayiLahir = True
On Error GoTo hell_
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoHasilPeriksa", adVarChar, adParamInput, 10, strNoHasilPeriksa)
        .Parameters.Append .CreateParameter("NoUrutBayi", adTinyInt, adParamInput, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adVarChar, adParamInput, 10, txtnopendaftaran.Text)
        If cbJK.Text = "Laki-Laki" Then
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, "L")
        Else
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, "P")
        End If
        .Parameters.Append .CreateParameter("NamaLengkapBayi", adVarChar, adParamInput, 40, IIf(txtNamaLengkap.Text = "", Null, txtNamaLengkap.Text))
        .Parameters.Append .CreateParameter("TinggiBeratBadan", adVarChar, adParamInput, 20, txtBeratBadan.Text + "/" + txtTinggiBadan.Text)
        .Parameters.Append .CreateParameter("WarnaKulit", adVarChar, adParamInput, 20, txtWarnaKulit.Text)
        .Parameters.Append .CreateParameter("KdKondisiLahir", adTinyInt, adParamInput, , dcKondisiBayi.BoundText)
        .Parameters.Append .CreateParameter("KdKuantitasLahirBayi", adChar, adParamInput, 1, dcKuantitas.BoundText)
        .Parameters.Append .CreateParameter("KdLetakJaninBayi", adTinyInt, adParamInput, , dcLetakJanin.BoundText)
        .Parameters.Append .CreateParameter("KdCaraLahirBayi", adTinyInt, adParamInput, , dcCaraLahir.BoundText)
        .Parameters.Append .CreateParameter("KdTempatLahirBayi", adTinyInt, adParamInput, , dcTempatLahirBayi.BoundText)
        .Parameters.Append .CreateParameter("Kelainan", adVarChar, adParamInput, 100, txtKelainan.Text)
        .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(dtpTglEventBayi.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoCmBayi", adChar, adParamInput, 10, IIf(txtnocm.Text = "", Null, txtnocm.Text))
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, txtKelainan.Text))
        .Parameters.Append .CreateParameter("KdPenyebabKematian", adTinyInt, adParamInput, , IIf(dcPenyebabKematian.BoundText = "", Null, dcPenyebabKematian.BoundText))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailHasilTindakanMedisBersalin"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan DATA", vbCritical, "Validasi"
            sp_AUD_DetailHasilTindakanMedisBersalin = False
        Else
            MsgBox "Penyimpanan data bayi lahir berhasil", vbInformation, "Validasi"
           
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    cmdSimpan.Enabled = False
Exit Function
hell_:
    msubPesanError ("AUD_DetailHasilTindakanMedisBersalin")
End Function

'Store procedure untuk mengisi registrasi pasien
Private Function sp_AddEventBayi() As Boolean
Set dbcmd = New ADODB.Command
sp_AddEventBayi = True
On Error GoTo hell_
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoUrutBayiLahir", adChar, adParamInput, 10, txtTempNourutBayi.Text)
        .Parameters.Append .CreateParameter("KdEvent", adChar, adParamInput, 3, dcNamaEvent.BoundText)
        .Parameters.Append .CreateParameter("TglEvent", adDate, adParamInput, , Format(dtpTglEventBayi.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
                        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_EventBayi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan DATA", vbCritical, "Validasi"
            sp_AddEventBayi = False
        Else
            MsgBox "Penyimpanan data event bayi lahir berhasil", vbInformation, "Validasi"
            Call Add_HistoryLoginActivity("Add_EventBayi")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell_:
    msubPesanError ("Add_EventBayi")
End Function

Private Sub cbJK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBeratBadan.SetFocus
End Sub

Private Sub cmdBayiLahir_Click()
On Error GoTo hell_
If dgDaftarPasienGD.ApproxCount = 0 Then Exit Sub
cmdSimpan.Enabled = True
    fraBayiLahir.Visible = True
    fraBayiLahir.Top = 1080
    fraBayiLahir.Left = 1680
    txtnopendaftaran.Text = dgDaftarPasienGD.Columns("NoPendaftaran").Value
    txtnocm.Text = dgDaftarPasienGD.Columns("NoCM").Value
    strNoHasilPeriksa = dgDaftarPasienGD.Columns("NoHasilPeriksa")
    
    cbJK.Text = ""
    txtBeratBadan.Text = ""
    txtTinggiBadan.Text = ""
    txtNamaLengkap.Text = ""
    txtWarnaKulit.Text = ""
    dcKondisiBayi.Text = ""
    dcPenyebabKematian.Text = ""
    dcPenyebabKematian.Enabled = False
    dcKuantitas.Text = ""
    dcLetakJanin.Text = ""
    dcCaraLahir.Text = ""
    dcTempatLahirBayi.Text = ""
    txtNoCMBayi.Text = ""
    txtKelainan.Text = ""
    txtKeterangan.Text = ""
    
    
   
Exit Sub
hell_:
    msubPesanError
End Sub

Public Sub cmdCari_Click()
On Error GoTo hell
'        mstrFilter = "AND TglMulaiPeriksa BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:ss") & "'"
'        strSQL = "SELECT TOP (200) [Nama Pasien], JK, Umur, JenisPasien, Kelas, SubInstalasi, UmurTahun, UmurBulan, UmurHari, KdJenisTarif, Alamat, NoHasilPeriksa, NoPakai, " & _
'                " NoLab_Rad, KeadaanLahirBayi, JenisTindakanMedis, KdKeadaanLahirBayi, KdJenisTindakanMedis, TglMulaiPeriksa, TglAkhirPeriksa, IdDokter," & _
'                " DokterPemeriksa , NoPendaftaran, NoCM,NoUrutBayi" & _
'                " FROM V_DaftarPasienVKBersalinBayiLahirx " & _
'                " where ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')  AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND KdRuangan='002' AND kdkelompokTM='04' or kdkelompokTM='14' or kdkelompokTM='19'  " & mstrFilter
        
        strSQL = "SELECT TOP (200) [Nama Pasien], JK, Umur, JenisPasien, Kelas, SubInstalasi, UmurTahun, UmurBulan, UmurHari, KdJenisTarif, Alamat, NoHasilPeriksa, NoPakai, " & _
                " NoLab_Rad, KeadaanLahirBayi, JenisTindakanMedis, KdKeadaanLahirBayi, KdJenisTindakanMedis, TglMulaiPeriksa, TglAkhirPeriksa, IdDokter," & _
                " DokterPemeriksa , NoPendaftaran, NoCM,NoUrutBayi" & _
                " FROM V_DaftarPasienVKBersalinBayiLahirx " & _
                " where TglMulaiPeriksa BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:ss") & "' AND ([Nama Pasien] like '%" & txtParameter.Text & "%' OR NoCM like '%" & txtParameter.Text & "%')  AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND KdRuangan='" & mstrKdRuangan & "' AND kdkelompokTM='04' or kdkelompokTM='14' or kdkelompokTM='19'  "

        
        Call msubRecFO(rs, strSQL)
        Set dgDaftarPasienGD.DataSource = rs
        
        With dgDaftarPasienGD
        .Columns("NoPendaftaran").Width = 1000
        .Columns("NoCM").Width = 800
        .Columns("Nama Pasien").Width = 2000
        .Columns("JK").Width = 400
        .Columns("Umur").Width = 1600
        .Columns("JenisPasien").Width = 1700
        .Columns("TglMulaiPeriksa").Width = 1900
        .Columns(TglAkhirPeriksa).Width = 1900
        .Columns("DokterPemeriksa").Width = 1500
        .Columns("SubInstalasi").Width = 1000
        .Columns("DokterPemeriksa").Width = 2000
        .Columns("SubInstalasi").Width = 1000
        .Columns("UmurTahun").Width = 600
        .Columns("UmurBulan").Width = 600
        .Columns("UmurHari").Width = 600
        .Columns("KdJenisTarif").Width = 0
        .Columns("Alamat").Width = 2500
        .Columns("IdDokter").Width = 2500
        .Columns("NoHasilPeriksa").Width = 1400
        .Columns("NoUrutBayi").Width = 1400
       
       End With
       
    lblJumData.Caption = "Data 0/" & rs.RecordCount
Exit Sub
hell:
    msubPesanError
End Sub

Private Sub cmdEventBayiLahir_Click()
On Error GoTo hell_
    If dgDaftarPasienGD.ApproxCount = 0 Then Exit Sub
    cmdSimpanEventBayi.Enabled = True
    fraEventBayi.Top = 2040
    fraEventBayi.Left = 2760
    If Len(dgDaftarPasienGD.Columns("NoUrutBayi")) = 0 Then
        MsgBox "Isi dulu data bayi lahir pada tombol Bayi Lahir", vbExclamation, "Validasi"
        cmdBayiLahir.SetFocus
        Exit Sub
    Else
        fraEventBayi.Visible = True
        txtNourutBayi.Text = dgDaftarPasienGD.Columns("NoUrutBayi")
        
    End If
Exit Sub
hell_:
    msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    If Periksa("combobox", cbJK, "Jenis Kelamin tidak boleh kosong!!!") = False Then Exit Sub

    If dcPenyebabKematian.Text <> "" Then
            If Periksa("datacombo", dcPenyebabKematian, "penyebab kematian salah") = False Then Exit Sub
    End If
    If Periksa("text", txtBeratBadan, "Berat badan tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("text", txtTinggiBadan, "Tinggi badan tidak boleh kosong!!!") = False Then Exit Sub
    
    If Periksa("text", txtWarnaKulit, "Warna kulit tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("datacombo", dcKondisiBayi, "Kondisi lahir tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("datacombo", dcKuantitas, "Kuantitas tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("datacombo", dcCaraLahir, "Cara lahir tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("datacombo", dcTempatLahirBayi, "Tempat lahir bayi tidak boleh kosong!!!") = False Then Exit Sub
    If Periksa("datacombo", dcLetakJanin, "Letak Janin tidak boleh kosong!!!") = False Then Exit Sub
    
    
    If sp_AUD_DetailHasilTindakanMedisBersalin("A") = False Then Exit Sub
    cmdSimpan.Enabled = False
    cmdTutupBayiLahir.SetFocus
End Sub

Private Sub cmdSimpanEventBayi_Click()
    cmdSimpanEventBayi.Enabled = False
    If Periksa("datacombo", dcNamaEvent, "Nama event bayi lahir tidak boleh kosong!!!") = False Then Exit Sub
    If sp_AddEventBayi() = False Then Exit Sub
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub cmdTutupBayiLahir_Click()
    Call cmdCari_Click
    fraBayiLahir.Visible = False
End Sub

Private Sub cmdTutupEventBayi_Click()
    cmdCari_Click
    fraEventBayi.Visible = False
End Sub


Private Sub dcCaraLahir_KeyDown(KeyCode As Integer, Shift As Integer)
   ' If KeyCode = 13 Then dcTempatLahirBayi.SetFocus
End Sub

Private Sub dcCaraLahir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

dcTempatLahirBayi.SetFocus
End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' and (JenisPasien LIKE '%" & dcJenisPasien.Text & "%')order by JenisPasien"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcJenisPasien.Text = ""
        dcKelas.SetFocus
        Exit Sub
        End If
        dcJenisPasien.BoundText = rs(0).Value
        dcJenisPasien.Text = rs(1).Value
    End If
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelas, DeskKelas FROM KelasPelayanan where StatusEnabled='1' and (DeskKelas LIKE '%" & DeskKelas.Text & "%')order by DeskKelas"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcKelas.Text = ""
        cmdCari.SetFocus
        Exit Sub
        End If
        dcKelas.BoundText = rs(0).Value
        dcKelas.Text = rs(1).Value
    End If
End Sub

Private Sub dcKondisiBayi_Change()
'      If dcKondisiBayi.BoundText = "" Then Exit Sub
'      If dcKondisiBayi.BoundText = 2 Then
'         dcPenyebabKematian.Enabled = True
'         dcPenyebabKematian.SetFocus
'       Else
'         dcPenyebabKematian.Text = ""
'         dcPenyebabKematian.Enabled = False
'         dcKuantitas.SetFocus
'       End If
End Sub

Private Sub dcKondisiBayi_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

dcKuantitas.SetFocus
End If
'     If dcKondisiBayi.BoundText = 2 Then
'        dcPenyebabKematian.Enabled = True
'        dcPenyebabKematian.SetFocus
'     Else
'        dcPenyebabKematian.Text = ""
'        dcPenyebabKematian.Enabled = False
'        dcKuantitas.SetFocus
'     End If

End Sub

Private Sub dcKondisiBayi_LostFocus()


 If dcKondisiBayi.BoundText = "" Then Exit Sub
      If dcKondisiBayi.BoundText = "2" Then
         dcPenyebabKematian.Enabled = True
         dcPenyebabKematian.SetFocus
       Else
         dcPenyebabKematian.Text = ""
         dcPenyebabKematian.Enabled = False
         dcKuantitas.SetFocus
       End If
End Sub

Private Sub dcKuantitas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

dcLetakJanin.SetFocus
End If
End Sub

Private Sub dcLetakJanin_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then dcCaraLahir.SetFocus
End Sub

Private Sub dcLetakJanin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

dcCaraLahir.SetFocus
End If
End Sub

Private Sub dcNamaEvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If dcNamaEvent.MatchedWithList = True Then dtpTglEventBayi.SetFocus
        strSQL = "SELECT  KdEvent, NamaEvent FROM EventBayi where StatusEnabled='1' and (NamaEvent LIKE '%" & dcNamaEvent.Text & "%')order by NamaEvent"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then
        dcNamaEvent.Text = ""
        dtpTglEventBayi.SetFocus
        Exit Sub
        End If
        dcNamaEvent.BoundText = rs(0).Value
        dcNamaEvent.Text = rs(1).Value
    End If
End Sub

Private Sub dcTempatLahirBayi_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then dtpTglLahirBayi.SetFocus
End Sub

Private Sub dgwDaftarPasienGD_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgDaftarPasienGD
        WheelHook.WheelHook dgDaftarPasienGD
End Sub

Private Sub dcTempatLahirBayi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

txtNoCMBayi.SetFocus
End If
End Sub

Private Sub dgDaftarPasienGD_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    lblJumData.Caption = "Data " & dgDaftarPasienGD.Bookmark & "/" & dgDaftarPasienGD.ApproxCount
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

Private Sub dtpTglEventBayi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSimpanEventBayi.SetFocus
End Sub

Private Sub dtpTglLahirBayi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoCMBayi.SetFocus
End Sub

Private Sub Form_Activate()
    cbJK.AddItem "Laki-Laki"
    cbJK.AddItem "Perempuan"
    Call cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strCtrlKey As String
Dim strShiftKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    strShiftKey = (Shift + vbShiftMask)
    Select Case KeyCode
        Case vbKeyF1

        Case vbKeyF9
            frmCtkDaftarPasienBersalin.Show
        Case vbKeyD

        Case vbKeyZ
            If strCtrlKey = 4 Then

            End If
               
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    dtpTglEventBayi.Value = Now
    dtpTglLahirBayi.Value = Now
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    
    'AddBy: Asep Nur Iman, untuk event bayi lahir di ruangan VK rawat inap
    Call msubRecFO(rsB, "Select value from SettingGlobal where prefix = 'kdRuanganVKRI'")
    If rsB.EOF = False Then
        If mstrKdRuangan = rsB(0).Value Then cmdBayiLahir.Visible = True
    End If
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienGD()
    With dgDaftarPasienGD
       .Columns("NoPendaftaran").Width = 1000
       .Columns("NoPendaftaran").Caption = "No. Registrasi"
       .Columns("NoCM").Width = 800
       .Columns("NoCM").Caption = "No.CM"
       .Columns("No.CM").Alignment = dbgCenter
       .Columns("Nama Pasien").Width = 2000
       .Columns("JK").Width = 400
       .Columns("Umur").Width = 1600
       .Columns("JenisPasien").Width = 1700
       .Columns("Kelas").Width = 1575
       .Columns("TglMulaiPeriksa").Width = 1900
       .Columns(TglAkhirPeriksa).Width = 1900
       .Columns("JenisPersalinan").Width = 1500
       .Columns("Dokter Pemeriksa").Width = 1500
       .Columns("SubInstalasi").Width = 1000
       .Columns("Dokter Pemeriksa").Width = 2000
       .Columns("SubInstalasi").Width = 1000
       .Columns("UmurTahun").Width = 600
       .Columns("UmurBulan").Width = 600
       .Columns("UmurHari").Width = 600
       .Columns("KdKelas").Width = 0
       .Columns("KdJenisTarif").Width = 0
       .Columns("Alamat").Width = 2500
       .Columns("IdDokter").Width = 2500
       .Columns("NoUrutBayiLahir").Width = 1400
       .Columns("NamaEvent").Width = 1400
       .Columns("KdEvent").Width = 0
       .Columns("TglEvent").Width = 0
       
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFormDaftarPasienIGD = False
End Sub

Private Sub txtBeratBadan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTinggiBadan.SetFocus
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtKelainan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub txtKondisiLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelainan.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaLengkap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtWarnaKulit.SetFocus
End Sub

Private Sub txtNoCMBayi_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then txtKelainan.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call cmdCari_Click
    txtParameter.SetFocus
    txtParameter.SelStart = Len(txtParameter.Text)
End Sub

'untuk load data pasien di form transaksi pelayanan
Private Sub subLoadFormTP()
On Error GoTo hell

    mstrNoPen = dgDaftarPasienGD.Columns("No. Registrasi").Value
    mstrNoCM = dgDaftarPasienGD.Columns("No. CM").Value
        
    If optRujukan.Value = True Then
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE (NamaLengkap = '" & dgDaftarPasienGD.Columns("Dokter Perujuk") & "')"
    Else
        strSQL = "SELECT IdPegawai FROM DataPegawai WHERE (NamaLengkap = '" & dgDaftarPasienGD.Columns("Dokter Penanggung Jawab") & "')"
    End If
    
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then mstrKdDokter = rs(0).Value Else mstrKdDokter = ""
    
    With frmTransaksiPasien
        .Show
        .txtnopendaftaran.Text = dgDaftarPasienGD.Columns("No. Registrasi").Value
        .txtnocm.Text = dgDaftarPasienGD.Columns("No. CM").Value
        .txtNamaPasien.Text = dgDaftarPasienGD.Columns("Nama Pasien").Value
        If dgDaftarPasienGD.Columns("JK").Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtKls.Text = dgDaftarPasienGD.Columns("Kelas").Value
        .txtThn.Text = dgDaftarPasienGD.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienGD.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienGD.Columns("UmurHari").Value
        
        If optRujukan.Value = True Then
            .txtJenisPasien.Text = dgDaftarPasienGD.Columns("JenisPasien")
            .txtTglDaftar.Text = dgDaftarPasienGD.Columns("TglPendaftaran").Value
            mdTglMasuk = dgDaftarPasienGD.Columns(16).Value
            mstrKdKelas = dgDaftarPasienGD.Columns(17).Value
            mstrKelas = dgDaftarPasienGD.Columns(18).Value
        Else
            .txtJenisPasien.Text = dgDaftarPasienGD.Columns("JenisPasien").Value
            .txtTglDaftar.Text = dgDaftarPasienGD.Columns(7).Value
            mdTglMasuk = dgDaftarPasienGD.Columns(7).Value
            mstrKdKelas = dgDaftarPasienGD.Columns("KdKelas").Value
        End If
        
        mstrKdSubInstalasi = dgDaftarPasienGD.Columns("KdSubInstalasi").Value
        
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    
    End With
Exit Sub
hell:
    Call msubPesanError
End Sub

'untuk load data pasien di form ubah jenis pasien
Private Sub subLoadFormJP()
On Error GoTo hell
    mstrNoPen = dgDaftarPasienGD.Columns("No. Registrasi").Value
    mstrNoCM = dgDaftarPasienGD.Columns("No. CM").Value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        .txtnocm.Text = dgDaftarPasienGD.Columns("No. CM").Value
        .txtNamaPasien.Text = dgDaftarPasienGD.Columns("Nama Pasien").Value
        If dgDaftarPasienGD.Columns("JK").Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgDaftarPasienGD.Columns("UmurTahun").Value
        .txtBln.Text = dgDaftarPasienGD.Columns("UmurBulan").Value
        .txtHr.Text = dgDaftarPasienGD.Columns("UmurHari").Value
        .dcJenisPasien.Text = dgDaftarPasienGD.Columns("JenisPasien").Value
        .lblNoPendaftaran.Visible = False
        .txtnopendaftaran.Visible = False
        mstrKdSubInstalasi = dgDaftarPasienGD.Columns("KdSubInstalasi").Value
        
        If optPasienNonRujukan.Value = True Then
            .txttglpendaftaran.Text = dgDaftarPasienGD.Columns("Tgl. Masuk").Value
        Else
            .txttglpendaftaran.Text = dgDaftarPasienGD.Columns("TglPendaftaran").Value
        End If
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
    End With
Exit Sub
hell:
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where StatusEnabled='1' order by JenisPasien")
    Call msubDcSource(dcNamaEvent, rs, "SELECT  KdEvent, NamaEvent FROM EventBayi where StatusEnabled='1' order by NamaEvent")
    Call msubDcSource(dcKondisiBayi, rs, "SELECT KdKeadaanLahirBayi, KeadaanLahirBayi FROM KeadaanLahirBayi where StatusEnabled='1' ")
    Call msubDcSource(dcKuantitas, rs, "SELECT KdKuantitasLahirBayi, KuantitasLahirBayi FROM KuantitasLahirBayi where StatusEnabled='1' ")
    Call msubDcSource(dcLetakJanin, rs, "SELECT  KdLetakJaninBayi, LetakJaninBayi FROM LetakJaninBayi where StatusEnabled='1' ")
    Call msubDcSource(dcCaraLahir, rs, "SELECT KdCaraLahirBayi, CaraLahirBayi FROM CaraLahirBayi where StatusEnabled='1' ")
    Call msubDcSource(dcTempatLahirBayi, rs, "SELECT KdTempatLahirBayi, TempatLahirBayi FROM TempatLahirBayi where StatusEnabled='1'")
    Call msubDcSource(dcPenyebabKematian, rs, "Select KdPenyebabKematian, PenyebabKematian from PenyebabKematian where StatusEnabled = '1'")
    
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtTinggiBadan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNamaLengkap.SetFocus
    ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWarnaKulit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKondisiBayi.SetFocus
End Sub
