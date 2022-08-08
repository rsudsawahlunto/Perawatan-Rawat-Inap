VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPemakaianObatAlkes2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemakaian Obat & Alkes"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemakaianObatAlkes2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   14340
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   960
   End
   Begin VB.PictureBox picPenerimaanSementara 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   2025
      ScaleWidth      =   7425
      TabIndex        =   44
      Top             =   2400
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame6 
         Caption         =   "Penerimaan Sementara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   7215
         Begin VB.CommandButton cmdSimpanTerimaBarang 
            Caption         =   "&Simpan"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1200
            Width           =   6735
         End
         Begin VB.TextBox txtNamaBarangPenerimaan 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            TabIndex        =   47
            Top             =   360
            Width           =   4815
         End
         Begin VB.TextBox txtJmlTerima 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   6960
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Terima Barang"
            Height          =   210
            Index           =   31
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Barang"
            Height          =   210
            Index           =   32
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   1065
         End
      End
   End
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   1920
      TabIndex        =   17
      Top             =   -1920
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4471
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
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   40
      Top             =   6720
      Width           =   14295
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12480
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
      Top             =   5760
      Width           =   14295
      Begin VB.TextBox txtTotalDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtTanggunganRS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9120
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtHutangPenjamin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtJumlahBayar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   32
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtHarusDibayar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11760
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtPembebasan 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   31
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   3840
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tanggungan RS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   9120
         TabIndex        =   38
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hutang Penjamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   6480
         TabIndex        =   37
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   3000
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Harus Dibayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   11760
         TabIndex        =   35
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pembebasan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   3240
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Biaya"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   945
      End
   End
   Begin MSDataGridLib.DataGrid dgDokter 
      Height          =   2295
      Left            =   11040
      TabIndex        =   15
      Top             =   -1680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
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
   Begin VB.Frame Frame8 
      Height          =   3855
      Left            =   0
      TabIndex        =   18
      Top             =   1920
      Width           =   14295
      Begin MSDataListLib.DataCombo dcNamaPelayananRS 
         Height          =   330
         Left            =   1560
         TabIndex        =   53
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.CheckBox chkStatusStok 
         Caption         =   "Ya"
         Height          =   495
         Left            =   3840
         TabIndex        =   51
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNoTemporary 
         Height          =   315
         Left            =   7080
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtHargaBeli 
         Height          =   315
         Left            =   4320
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtJenisBarang 
         Height          =   315
         Left            =   3000
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtKdDokter 
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAsalBarang 
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdAsal 
         Height          =   315
         Left            =   2760
         TabIndex        =   21
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSatuan 
         Height          =   315
         Left            =   4920
         TabIndex        =   20
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtKdBarang 
         Height          =   315
         Left            =   3720
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcJenisObat 
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Resep"
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
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   14295
      Begin VB.TextBox txtRP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   11520
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   7560
         TabIndex        =   5
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtNoResep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkDokterPemeriksa 
         Caption         =   "Dokter Penulis Resep"
         Height          =   255
         Left            =   7560
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkNoResep 
         Caption         =   "No. Resep"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpTglPelayanan 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   127074307
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpTglResep 
         Height          =   330
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   127074307
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Perawatan"
         Height          =   210
         Index           =   10
         Left            =   11520
         TabIndex        =   41
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Resep"
         Height          =   210
         Index           =   1
         Left            =   5880
         TabIndex        =   24
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   43
      Top             =   7605
      Visible         =   0   'False
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   25250
            MinWidth        =   11359
            Text            =   "Daftar Pemakain Barang Gratis (F5)"
            TextSave        =   "Daftar Pemakain Barang Gratis (F5)"
         EndProperty
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
      TabIndex        =   52
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
      Left            =   12480
      Picture         =   "frmPemakaianObatAlkes2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPemakaianObatAlkes2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemakaianObatAlkes2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmPemakaianObatAlkes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim subintJmlArray As Integer
Dim subcurHargaSatuan As Currency
Dim subcurTarifService As Currency
Dim subcurHarusDibayar As Currency
Dim curTanggunganRS As Currency
Dim curHutangPenjamin As Currency
Dim subintJmlService As Integer
Dim tempStatusTampil As Boolean
Dim subJenisHargaNetto As Integer
Dim StrKdRP As String

Public Function sp_StokRuangan(f_KdBarang As String, f_KdAsal As String, f_JmlBarang As Double, f_status As String) As Boolean
    On Error GoTo errLoad
    Dim i As Integer

    sp_StokRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , f_JmlBarang)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_StokRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
            sp_StokRuangan = False
        Else
            Call Add_HistoryLoginActivity("Update_StokRuangan")
        End If
    End With
    Set dbcmd = Nothing

    Exit Function
errLoad:
    sp_StokRuangan = False
    Call msubPesanError
End Function

Private Function sp_GenerateNoResep() As Boolean
    On Error GoTo errLoad

    sp_GenerateNoResep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, Trim(txtNoResep.Text))
        .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("OutputNoResep", adVarChar, adParamOutput, 15, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_GenerateNoResep"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_GenerateNoResep = False
        Else
            txtNoResep.Text = Trim(.Parameters("OutputNoResep"))

        End If
    End With

    Exit Function
errLoad:
    sp_GenerateNoResep = False
    Call msubPesanError("sp_GenerateNoResep")
End Function

Private Sub chkDokterPemeriksa_Click()
'    On Error GoTo errLoad
'
'    If chkDokterPemeriksa.Value = vbUnchecked Then
'        txtDokter.Enabled = False
'        txtDokter.Text = ""
'    Else
'        txtDokter.Enabled = True
'        txtDokter.Text = mstrNamaDokter
'        txtKdDokter.Text = mstrKdDokter
'    End If
'    dgDokter.Visible = False
'
'    Exit Sub
'errLoad:
'    Call msubPesanError

    On Error GoTo errLoad

    If chkDokterPemeriksa.Value = 0 Then
        txtDokter.Enabled = False
'        txtDokter.Text = ""

        If dgDokter.Visible = True Then dgDokter.Visible = False
    Else
        
        txtDokter.Enabled = True
        strSQL = "SELECT dbo.RegistrasiRI.IdDokter, dbo.DataPegawai.NamaLengkap " & _
        " FROM dbo.RegistrasiRI INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRI.IdDokter = dbo.DataPegawai.IdPegawai " & _
        " WHERE (dbo.RegistrasiRI.NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)

        If Not rs.EOF Then
            txtDokter.Text = rs(1).Value
            dgDokter.Visible = False
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDokterPemeriksa_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If chkDokterPemeriksa.Value = vbChecked Then txtDokter.SetFocus Else fgData.SetFocus
    End If
End Sub

Private Sub chkNoResep_Click()
    If chkNoResep.Value = vbChecked Then
        txtNoResep.Enabled = True
        dtpTglResep.Enabled = True
        chkDokterPemeriksa.Enabled = True
        txtDokter.Enabled = True
        'txtResepKe.Enabled = True
    Else
        txtNoResep.Enabled = False
        dtpTglResep.Enabled = False
        chkDokterPemeriksa.Enabled = False
        txtDokter.Enabled = False
        chkDokterPemeriksa.Value = vbUnchecked
        'txtResepKe.Enabled = False
    End If
    dgDokter.Visible = False
End Sub

Private Sub chkNoResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNoResep.Enabled = True Then txtNoResep.SetFocus Else fgData.SetFocus
    End If
End Sub

Private Sub chkStatusStok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkStatusStok.Visible = False
        fgData.TextMatrix(fgData.Row, fgData.Col) = IIf(chkStatusStok.Value = vbChecked, "Ya", "Tdk")
        With fgData
            If .RowPos(.Row) >= .Height - 360 Then
                .SetFocus
                SendKeys "{DOWN}"
                Exit Sub
            End If
            .SetFocus

            If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then
                fgData.Rows = fgData.Rows + 1
                If .TextMatrix(.Rows - 2, 25) = "" Then
                    .TextMatrix(.Rows - 1, 0) = "1"
                ElseIf .TextMatrix(.Rows - 2, 25) = "01" Then
                    .TextMatrix(.Rows - 1, 0) = "1"
                Else
                    .TextMatrix(.Rows - 1, 0) = Val(.TextMatrix(.Rows - 2, 0))
                End If
            End If

            fgData.SetFocus
            fgData.Row = fgData.Rows - 1
            fgData.Col = 0
        End With
    End If
End Sub

Private Sub chkStatusStok_LostFocus()
    chkStatusStok.Visible = False
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer

    If txtDokter.Text = "" Then MsgBox "Dokter penulis resep harus diisi", vbExclamation, "Validasi": Exit Sub
'    If dcJenisObat.BoundText = "" Then MsgBox "Data barang harus diisi lengkap", vbExclamation, "Validasi": Exit Sub
'    If fgData.TextMatrix(1, 2) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub
    If fgData.TextMatrix(1, 1) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub
    If fgData.TextMatrix(1, 10) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub

    If sp_GenerateNoResep() = False Then Exit Sub
    If sp_ResepObat() = False Then Exit Sub

    With fgData
        For i = 1 To .Rows - 1
            dtpTglPelayanan.Value = Now
            If .TextMatrix(i, 2) = "" Then GoTo lanjut_
            Set dbRst = Nothing
              strSQL = "SELECT * FROM PemakaianAlkes where NoPendaftaran ='" & mstrNoPen & "' and KdBarang like '%" & .TextMatrix(i, 2) & "%' and Kdasal like '%" & .TextMatrix(i, 12) & "%' and TglPelayanan ='" & Format(dtpTglPelayanan.Value, "yyyy/mm/dd hh:mm:ss") & "'"
              Call msubRecFO(dbRst, strSQL)
              If dbRst.EOF = False Then dtpTglPelayanan.Value = DateAdd("s", 1, dtpTglPelayanan.Value)
            
            If sp_PemakaianObatAlkesResep(.TextMatrix(i, 2), .TextMatrix(i, 12), .TextMatrix(i, 6), _
                CDbl(.TextMatrix(i, 10)), .TextMatrix(i, 16), .TextMatrix(i, 25), .TextMatrix(i, 15), .TextMatrix(i, 14), .TextMatrix(i, 0), IIf(LCase(.TextMatrix(i, 27)) = "ya", "1", "0"), .TextMatrix(i, 29), "", "", "", "", "", .TextMatrix(i, 31), dtpTglPelayanan.Value) = False Then Exit Sub

                If .TextMatrix(i, 30) = "1" Then If update_DetailOrderTMOA(dbcmd, fgData.TextMatrix(i, 2), "OA") = False Then Exit Sub

            Next i
        End With
        dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"

lanjut_:
        MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
        Call Add_HistoryLoginActivity("AU_GenerateNoResep+Add_ResepObat+Add_PemakaianObatAlkesResep")
        txtNoResep.Text = ""
        txtTotalBiaya.Text = 0
        txtTotalDiscount.Text = 0
        txtHutangPenjamin.Text = 0
        txtTanggunganRS.Text = 0
        txtHarusDibayar.Text = 0
        cmdSimpan.Enabled = False

        Exit Sub
errLoad:
        Call msubPesanError
End Sub

'Store procedure untuk mengisi biaya pelayanan pasien
'-----------------------**--Yang ditambah--**-----------------------
Private Function update_DetailOrderTMOA(ByVal adoCommand As ADODB.Command, sItem As String, sStatus As String) As Boolean
    On Error GoTo errLoad
    update_DetailOrderTMOA = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdItem", adVarChar, adParamInput, 9, sItem)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, sStatus)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_DetailOrderTMOA"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            update_DetailOrderTMOA = False

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

    Exit Function
errLoad:
    update_DetailOrderTMOA = False
    Call msubPesanError
End Function

Private Sub cmdSimpanTerimaBarang_Click()
    On Error GoTo errLoad

    If Val(txtJmlTerima.Text) = 0 Then Exit Sub
    If sp_PenerimaanSementara(Now, fgData.TextMatrix(fgData.Row, 2), fgData.TextMatrix(fgData.Row, 12), Val(txtJmlTerima.Text), "A") = False Then Exit Sub
    Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & fgData.TextMatrix(fgData.Row, 2) & "','" & fgData.TextMatrix(fgData.Row, 12) & "') as stok")
    fgData.TextMatrix(fgData.Row, 9) = rs(0)
    picPenerimaanSementara.Visible = False
    fgData.SetFocus: fgData.Col = 10

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    'Dim i As Integer
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data Pemakaian Obat dan Alat Kesehatan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If

    End If
    dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & txtNoTemporary & "')"
    Unload Me
    Call frmTransaksiPasien.subPemakaianObatAlkes
End Sub

Private Sub dcJenisObat_Change()
    On Error GoTo errLoad
    Dim i As Integer

    subcurTarifService = 0

    If bolStatusFIFO = False Then
        fgData.TextMatrix(fgData.Row, 1) = dcJenisObat.Text
        fgData.TextMatrix(fgData.Row, 25) = dcJenisObat.BoundText
        fgData.TextMatrix(fgData.Row, 14) = subcurTarifService
    Else
        With fgData
            For i = 1 To .Rows - 1
                If .TextMatrix(.Row, 2) = .TextMatrix(i, 2) And .TextMatrix(.Row, 12) = .TextMatrix(i, 12) And .TextMatrix(.Row, 6) = .TextMatrix(i, 6) Then
                    .TextMatrix(i, 1) = dcJenisObat.Text
                    .TextMatrix(i, 25) = dcJenisObat.BoundText
                    .TextMatrix(i, 14) = subcurTarifService
                End If
            Next i
        End With
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisObat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then dcJenisObat.Visible = False: fgData.SetFocus
End Sub

Private Sub dcJenisObat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcJenisObat_Change
        dcJenisObat.Visible = False
        fgData.Col = 3
        fgData.SetFocus
    ElseIf KeyAscii = 27 Then
        dcJenisObat.Visible = False
    End If
End Sub

Private Sub dcJenisObat_LostFocus()
    dcJenisObat.Visible = False
End Sub

Private Sub dcNamaPelayananRS_Change()
    On Error GoTo errLoad
    Dim i As Integer

    If bolStatusFIFO = False Then
        fgData.TextMatrix(fgData.Row, 28) = dcNamaPelayananRS.Text
        fgData.TextMatrix(fgData.Row, 29) = dcNamaPelayananRS.BoundText
    Else
        With fgData
            For i = 1 To .Rows - 1
                If .TextMatrix(.Row, 2) = .TextMatrix(i, 2) And .TextMatrix(.Row, 12) = .TextMatrix(i, 12) And .TextMatrix(.Row, 6) = .TextMatrix(i, 6) Then
                    .TextMatrix(i, 28) = dcNamaPelayananRS.Text
                    .TextMatrix(i, 29) = dcNamaPelayananRS.BoundText
                End If
            Next i
        End With
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNamaPelayananRS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcNamaPelayananRS_Change
        dcNamaPelayananRS.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcNamaPelayananRS_LostFocus()
    dcNamaPelayananRS.Visible = False
End Sub

Private Sub dgDokter_DblClick()
    On Error GoTo errLoad
    If dgDokter.ApproxCount = 0 Then Exit Sub
    txtDokter.Text = dgDokter.Columns("Nama Dokter")
    dgDokter.Visible = False
    txtKdDokter.Text = dgDokter.Columns("KodeDokter")
    fgData.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgDokter_DblClick
End Sub

Private Sub dgObatAlkes_DblClick()
    On Error GoTo errLoad
    Dim i As Integer
    Dim tempSettingDataPendukung As Integer
    Dim curHargaBrg As Currency
    Dim strNoTerima As String

    curHutangPenjamin = 0
    curTanggunganRS = 0

    For i = 0 To fgData.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgData.TextMatrix(i, 2) And dgObatAlkes.Columns("KdAsal") = fgData.TextMatrix(i, 12) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgData.SetFocus: fgData.Row = i
            Exit Sub
        End If
    Next i

    strNoTerima = ""

    Set rsB = Nothing
    Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
    strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))

    For i = 0 To fgData.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgData.TextMatrix(i, 2) And dgObatAlkes.Columns("KdAsal") = fgData.TextMatrix(i, 12) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgData.SetFocus: fgData.Row = i
            Exit Sub
        End If
    Next i
    
    strNoTerima = ""

    Set rsB = Nothing
    Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & mstrKdRuangan & "') as NoFIFO")
    strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))

    For i = 0 To fgData.Rows - 1
        If dgObatAlkes.Columns("KdBarang") = fgData.TextMatrix(i, 2) And dgObatAlkes.Columns("KdAsal") = fgData.TextMatrix(i, 12) Then
            MsgBox "Data tersebut sudah diinput", vbExclamation, "Validasi"
            dgObatAlkes.Visible = False
            fgData.SetFocus: fgData.Row = i
            Exit Sub
        End If
    Next i

    With fgData
        .TextMatrix(.Row, 2) = dgObatAlkes.Columns("KdBarang")
        .TextMatrix(.Row, 3) = dgObatAlkes.Columns("NamaBarang")
        .TextMatrix(.Row, 4) = dgObatAlkes.Columns("Kekuatan")
        .TextMatrix(.Row, 5) = dgObatAlkes.Columns("AsalBarang")
        .TextMatrix(.Row, 6) = dgObatAlkes.Columns("Satuan")
        '.TextMatrix(.Row, 7) = Format(dgObatAlkes.Columns("HargaBarang").Value, "#,###")
        .TextMatrix(.Row, 31) = strNoTerima
        curHargaBrg = 0

        strSQL = ""
        Set rsB = Nothing
        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "','" & dgObatAlkes.Columns("Satuan") & "', '" & mstrKdRuangan & "','" & .TextMatrix(.Row, 31) & "') AS HargaBarang"
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

        strSQL = ""
        Set rs = Nothing
        subcurHargaSatuan = 0
        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & dgObatAlkes.Columns("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
        Call msubRecFO(rs, strSQL)
        'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
'        .TextMatrix(.Row, 7) = subcurHargaSatuan
'        .TextMatrix(.Row, 7) = Format(subcurHargaSatuan, "#,###")
        .TextMatrix(.Row, 7) = IIf(subcurHargaSatuan = 0, 0, subcurHargaSatuan)
        .TextMatrix(.Row, 7) = IIf(Format(subcurHargaSatuan, "#,###") = "", 0, Format(subcurHargaSatuan, "#,###"))
        .TextMatrix(.Row, 8) = (dgObatAlkes.Columns("Discount").Value / 100) * subcurHargaSatuan
        
'        strSQL = ""
'        Set rs = Nothing
'        strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & dgObatAlkes.Columns("KdBarang") & "' and KdAsal='" & dgObatAlkes.Columns("KdAsal") & "' and KdRuangan='" & mstrKdRuangan & "'"
'        Call msubRecFO(rs, strSQL)
'        If rs.EOF Then
'            .TextMatrix(.Row, 9) = 0
'        Else
'            .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))
'        End If

        Call msubRecFO(rs, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & dgObatAlkes.Columns("KdBarang") & "','" & dgObatAlkes.Columns("KdAsal") & "') as stok")
        .TextMatrix(.Row, 9) = IIf(IsNull(rs("Stok")), 0, rs("Stok"))

        .TextMatrix(.Row, 12) = dgObatAlkes.Columns("KdAsal")
        .TextMatrix(.Row, 13) = dgObatAlkes.Columns("JenisBarang")
        .TextMatrix(.Row, 14) = subcurTarifService
        .TextMatrix(.Row, 15) = subintJmlService
        .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))
        .TextMatrix(.Row, 17) = curHutangPenjamin
        .TextMatrix(.Row, 18) = curTanggunganRS
        .TextMatrix(.Row, 19) = 0
        .TextMatrix(.Row, 20) = 0
        .TextMatrix(.Row, 21) = 0

        .TextMatrix(.Row, 23) = txtNoTemporary.Text
        txtHargaBeli.Text = curHargaBrg 'dgObatAlkes.Columns("HargaBarang")
        .TextMatrix(.Row, 24) = CDbl(txtHargaBeli.Text)

    End With

    dgObatAlkes.Visible = False
    txtJenisBarang.Text = "": txtKdBarang.Text = "": txtKdAsal.Text = "": txtSatuan.Text = "": txtAsalBarang.Text = "": 'txtKekuatan.Text = ""

    With fgData
        .SetFocus
        If .Col = 2 Then
            .Col = 3
        ElseIf .Col = 3 Then
            .Col = 10
        End If
    End With

    Exit Sub
errLoad:
End Sub

Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgObatAlkes_DblClick
End Sub

Private Sub dtpTglPelayanan_Change()
    dtpTglPelayanan.MaxDate = Now
End Sub

Private Sub dtpTglPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then fgData.SetFocus
End Sub

Private Sub dtpTglResep_Change()
    dtpTglResep.MaxDate = Now
End Sub

Private Sub dtpTglResep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkDokterPemeriksa.SetFocus
End Sub

Private Sub dtpTglResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkDokterPemeriksa.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKdBrg As String
    Dim strKdAsal As String
    Dim i As Integer

    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then
                        fgData.Rows = fgData.Rows + 1
                        If fgData.TextMatrix(fgData.Rows - 2, 25) = "" Then
                            fgData.TextMatrix(fgData.Rows - 1, 0) = "1"
                        ElseIf fgData.TextMatrix(fgData.Rows - 2, 25) = "01" Then
                            fgData.TextMatrix(fgData.Rows - 1, 0) = "0"
                        Else
                            fgData.TextMatrix(fgData.Rows - 1, 0) = Val(fgData.TextMatrix(fgData.Rows - 2, 0))
                        End If
                    End If
                    fgData.Row = fgData.Rows - 1
                    fgData.Col = 1
                Else
                    fgData.Col = 1
                End If
            Else
                For i = 0 To fgData.Cols - 2
                    If fgData.Col = fgData.Cols - 1 Then Exit For
                    fgData.Col = fgData.Col + 1
                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                Next i
            End If
            fgData.SetFocus
            If fgData.Col = 1 Then Call subLoadDataCombo(dcJenisObat)

        Case 27
            dgObatAlkes.Visible = False

        Case vbKeyDelete
            'validasi FIFO
            If bolStatusFIFO = True Then
                If fgData.CellBackColor = vbRed Then
                    MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
                    fgData.SetFocus
                    Exit Sub
                End If

                With fgData
                    'If Trim(.TextMatrix(.Row, 10)) <> "" Then
                    i = .Rows - 1
                    strKdBrg = .TextMatrix(.Row, 2)
                    strKdAsal = .TextMatrix(.Row, 12)
                    Do While i <> 0 'khusus utk delete dr keyboard diset 0 agar ke cek keseluruhannya
                        If .TextMatrix(i, 2) <> "" Then
                            If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                .Row = i
                                Call subHapusDataGrid
'                                .Row = i - 1
                            End If
                        End If
                        i = i - 1
                    Loop
                    
                    'End If
                End With
            Else
                Call subHapusDataGrid
            End If
            'end FIFO
    End Select
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    'Validasi jika FIFO
    If bolStatusFIFO = True Then
        If fgData.CellBackColor = vbRed Then
            MsgBox "Data yang barisnya berwarna merah tidak bisa di edit", vbExclamation, "validasi"
            fgData.SetFocus
            Exit Sub
        End If
    End If
    'end fifo

    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgData.Col
        Case 0 'R/Ke

            txtIsi.MaxLength = 2
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 1 'Jenis Obat
            fgData.Col = 1
            Call subLoadDataCombo(dcJenisObat)

        Case 2 'Kode Barang
            txtIsi.MaxLength = 9
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 3 'Nama Barang
            txtIsi.MaxLength = 20
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 10 'Jumlah
            txtIsi.MaxLength = 7
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Or KeyAscii = Asc(".")) Then Exit Sub
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

        Case 27 'Status Stok
            Call subLoadCheck

        Case 28 'nama pelayanan rs yang di gunakan ' ganti pemakain bahan
            fgData.Col = 28
            Call subLoadDataCombo(dcNamaPelayananRS)
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            'If cmdSimpan.Enabled = False Then frmDaftarBarangGratisRuangan.Show
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Dim curHargaBrg As Currency
    Dim i, j, k, intRowTemp, iTemp As Integer
    Dim tempSettingDataPendukung As Integer
    Dim curHarusDibayar As Currency
    Dim strNoTerima As String
    Dim strKdBrg, strKdAsal As String
    Dim dblSelisih As Double
    Dim rsC As ADODB.recordset
    Dim bolCekFIFO As Boolean

    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglPelayanan.Value = Now
    dtpTglResep.Value = Now

    Call subSetGrid
    Call subLoadDcSource
    dgDokter.Visible = False
    dcJenisObat.BoundText = ""

    dgObatAlkes.Top = 2880
    dgObatAlkes.Left = 2040
    dgObatAlkes.Visible = True
    dgObatAlkes.Visible = False

    strSQL = "SELECT JenisHargaNetto" & _
    " From PersentaseUpTarifOA" & _
    " Where(IdPenjamin = '" & mstrKdPenjaminPasien & "') And (KdKelompokPasien = '" & mstrKdJenisPasien & "')"
    Call msubRecFO(rs, strSQL)
    subJenisHargaNetto = IIf(rs.EOF = True, 1, rs(0))

    ' untuk insert otomatis jika ada order
    bolCekFIFO = False
    strSQL = "SELECT * FROM V_DetailOrderOAx where NoPendaftaran ='" & mstrNoPen & "' and KdRuanganTujuan ='" & mstrKdRuangan & "' and KdRuanganStok='" & mstrKdRuangan & "'"
    Set dbRst = Nothing
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = False Then
        If IsNull(dbRst("IdDokterOrder")) Then
            chkDokterPemeriksa.Value = Unchecked
        Else
            chkDokterPemeriksa.Value = Checked
            txtDokter.Text = dbRst("DokterOrder")
            txtKdDokter.Text = dbRst("IdDokterOrder")
            dgDokter.Visible = False
        End If
        With fgData
            For i = 1 To dbRst.RecordCount
                curHutangPenjamin = 0
                curTanggunganRS = 0

                If bolStatusFIFO = True Then
                    iTemp = i
                    If bolCekFIFO = True Then i = i + 1
                End If

                .TextMatrix(i, 0) = dbRst("ResepKe")
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 25) = ""
                .TextMatrix(i, 2) = dbRst("KdBarang")
                .TextMatrix(i, 3) = dbRst("NamaBarang")
                .TextMatrix(i, 4) = IIf(IsNull(dbRst("Kekuatan")), "", dbRst("Kekuatan"))
                .TextMatrix(i, 5) = dbRst("NamaAsal")
                .TextMatrix(i, 6) = dbRst("Satuan")
                .TextMatrix(i, 12) = dbRst("KdAsal")

                Set rsB = Nothing
                Call msubRecFO(rsB, "select dbo.TakeNoFIFO_F('" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 12) & "','" & mstrKdRuangan & "') as NoFIFO")
                strNoTerima = IIf(IsNull(rsB("NoFIFO")), "0000000000", rsB("NoFIFO"))
                .TextMatrix(i, 31) = strNoTerima

                curHargaBrg = 0
                strSQL = ""
                Set rsB = Nothing
                strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & dbRst("KdBarang") & "','" & dbRst("KdAsal") & "','" & dbRst("Satuan") & "', '" & mstrKdRuangan & "', '" & .TextMatrix(i, 31) & "') AS HargaBarang"
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value
                strSQL = ""
                Set rsB = Nothing
                subcurHargaSatuan = 0
                strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & dbRst("KdAsal") & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rsB(0).Value
                .TextMatrix(i, 7) = subcurHargaSatuan
                .TextMatrix(i, 7) = Format(subcurHargaSatuan, "#,###")
                .TextMatrix(i, 8) = (0 / 100) * subcurHargaSatuan ' diskon di set 0
                strSQL = ""
                Set rsB = Nothing
                strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & .TextMatrix(i, 2) & "' and KdAsal='" & .TextMatrix(i, 12) & "' and KdRuangan='" & mstrKdRuangan & "'"
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF Then
                    .TextMatrix(i, 9) = 0
                Else
                    .TextMatrix(i, 9) = IIf(IsNull(rsB("Stok")), 0, rsB("Stok"))
                End If

                .TextMatrix(i, 10) = dbRst("JmlBarang")

                'perhitungan jika JmlBarang lebih dari penerimaan
                'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                If bolStatusFIFO = True Then
                    If Trim(.TextMatrix(i, 10)) <> "" Then
                        j = .Rows - 1
                        strKdBrg = .TextMatrix(i, 2)
                        strKdAsal = .TextMatrix(i, 12)
                        Do While j <> 1
                            If .TextMatrix(j, 2) <> "" Then
                                If (strKdBrg = .TextMatrix(j, 2)) And (strKdAsal = .TextMatrix(j, 12)) Then
                                    .Row = j
                                    If .CellBackColor = vbRed Then
                                        Call subHapusDataGrid
                                        .Row = j - 1
                                    End If
                                End If
                            End If
                            j = j - 1
                        Loop

                        For j = 1 To .Rows - 1
                            If (strKdBrg = .TextMatrix(j, 2)) And (strKdAsal = .TextMatrix(j, 12)) Then
                                .Row = j
                                Exit For
                            End If
                        Next j
                    End If

                    intRowTemp = 0
                End If

                'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                If bolStatusFIFO = True Then
                    If Trim(.TextMatrix(i, 10)) = "" Then .TextMatrix(i, 10) = 0
                    Set rsB = Nothing
                    Call msubRecFO(rsB, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 12) & "','" & .TextMatrix(i, 31) & "') as stok")
                    dblSelisih = rsB(0) - CDbl(.TextMatrix(i, 10))
                    If dblSelisih < 0 Then
                        .TextMatrix(i, 10) = rsB(0)
                        '.TextMatrix(.Row, 9) = dbRst(0)
                    Else
                        Set rsB = Nothing
                        strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & .TextMatrix(i, 2) & "' and KdAsal='" & .TextMatrix(i, 12) & "' and KdRuangan='" & mstrKdRuangan & "'"
                        Call msubRecFO(rsB, strSQL)
                        If rsB.EOF Then
                            .TextMatrix(i, 9) = 0
                        Else
                            .TextMatrix(i, 9) = IIf(IsNull(rsB("Stok")), 0, rsB("Stok"))
                        End If
                    End If
                End If
                'end FIFO

                .TextMatrix(i, 13) = dbRst("JenisBarang")

                .TextMatrix(i, 14) = 1
                .TextMatrix(i, 15) = 1

                If sp_TempDetailApotikJual(CDbl(.TextMatrix(i, 7)) + CDbl(.TextMatrix(i, 14)), .TextMatrix(i, 2), .TextMatrix(i, 12)) = False Then Exit Sub
                strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
                " FROM TempDetailApotikJual" & _
                " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(i, 2) & "') AND (KdAsal = '" & .TextMatrix(i, 12) & "')"
                Set rsB = Nothing
                Call msubRecFO(rsB, strSQL)
                If rsB.EOF = True Then
                    curHutangPenjamin = 0
                    curTanggunganRS = 0
                Else
                    curHutangPenjamin = rsB("JmlHutangPenjamin").Value
                    curTanggunganRS = rsB("JmlTanggunganRS").Value
                End If

                .TextMatrix(i, 16) = CDbl(.TextMatrix(i, 7))
                .TextMatrix(i, 17) = curHutangPenjamin
                .TextMatrix(i, 18) = curTanggunganRS
                .TextMatrix(i, 19) = 0
                .TextMatrix(i, 20) = 0
                .TextMatrix(i, 21) = 0

                .TextMatrix(i, 23) = txtNoTemporary.Text
                txtHargaBeli.Text = curHargaBrg 'dgObatAlkes.Columns("HargaBarang")
                .TextMatrix(i, 24) = CDbl(txtHargaBeli.Text)

                tempSettingDataPendukung = 0
                '                For j = 1 To .Rows - 1
                '                    .TextMatrix(j, 26) = 0
                '                    If i = (typSettingDataPendukung.intJumlahBAdminOAPerBaris * tempSettingDataPendukung) + 1 Then
                '                        tempSettingDataPendukung = tempSettingDataPendukung + 1
                '                        .TextMatrix(j, 26) = typSettingDataPendukung.curBiayaAdministrasi
                '                    End If
                '                Next j
                .TextMatrix(i, 26) = 0

                .TextMatrix(i, 11) = ((CDbl(.TextMatrix(i, 14)) * CDbl(.TextMatrix(i, 15))) + _
                (CDbl(.TextMatrix(i, 16)) * Val(.TextMatrix(i, 10)))) + Val(.TextMatrix(i, 26))

                If curHutangPenjamin > 0 Then
                    .TextMatrix(i, 19) = (.TextMatrix(i, 14) * .TextMatrix(i, 15)) + (CDbl(.TextMatrix(i, 10)) * CDbl(.TextMatrix(i, 17))) + Val(.TextMatrix(i, 26))
                Else
                    .TextMatrix(i, 19) = 0
                End If

                If curTanggunganRS > 0 Then
                    .TextMatrix(i, 20) = (.TextMatrix(i, 14) * .TextMatrix(i, 15)) + (CDbl(.TextMatrix(i, 10)) * CDbl(.TextMatrix(i, 18))) + Val(.TextMatrix(i, 26))
                Else
                    .TextMatrix(i, 20) = 0
                End If
                .TextMatrix(i, 21) = CDbl(.TextMatrix(i, 10)) * CDbl(.TextMatrix(i, 8))

                'total harus dibayar = total harga - total discount - _
                total hutang penjamin - totaltanggunganrs
                curHarusDibayar = CDbl(.TextMatrix(i, 11)) - (CDbl(.TextMatrix(i, 21)) + _
                CDbl(.TextMatrix(i, 19)) + CDbl(.TextMatrix(i, 20)))
                .TextMatrix(i, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                .TextMatrix(i, 27) = 0
                .TextMatrix(i, 28) = IIf(IsNull(dbRst("NamaPelayanan")), "", dbRst("NamaPelayanan"))
                .TextMatrix(i, 29) = IIf(IsNull(dbRst("KdPelayananRSUsed")), "", dbRst("KdPelayananRSUsed"))
                .TextMatrix(i, 30) = "1"

                'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                If bolStatusFIFO = True Then
                    If dblSelisih < 0 Then
                        bolCekFIFO = True
                        With fgData
                            strSQL = "select NoTerima As NoFIFO,JmlStokMax from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 31) & "' and JmlStok<>0 order by TglTerima asc"
                            Set rsC = Nothing
                            Call msubRecFO(rsC, strSQL)
                            If rsC.EOF = False Then
                                rsC.MoveFirst
                                For k = 1 To rsC.RecordCount
                                    '.Rows = .Rows - 1

                                    .Rows = .Rows + 1
                                    i = i + 1

                                    intRowTemp = .Row
                                    If .TextMatrix(.Rows - 2, 2) = "" Then
                                        .Row = .Rows - 2
                                    Else
                                        .Row = .Rows - 1
                                    End If
                                    For j = 0 To .Cols - 1
                                        .Col = j
                                        .CellBackColor = vbRed
                                        .CellForeColor = vbWhite
                                    Next j

                                    .Row = intRowTemp
                                    intRowTemp = 0
                                    If .TextMatrix(.Rows - 2, 2) = "" Then
                                        intRowTemp = .Rows - 2
                                    Else
                                        intRowTemp = .Rows - 1
                                    End If

                                    curHutangPenjamin = 0
                                    curTanggunganRS = 0

                                    .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                                    .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                                    .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                                    .TextMatrix(intRowTemp, 3) = .TextMatrix(.Row, 3)
                                    .TextMatrix(intRowTemp, 4) = .TextMatrix(.Row, 4)
                                    .TextMatrix(intRowTemp, 5) = .TextMatrix(.Row, 5)
                                    .TextMatrix(intRowTemp, 6) = .TextMatrix(.Row, 6)
                                    .TextMatrix(intRowTemp, 12) = .TextMatrix(.Row, 12)

                                    '                        Call msubRecFO(rsb, "select dbo.TakeNoFIFO_F('" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 12) & "','" & mstrKdRuangan & "') as NoFIFO")
                                    '                        strNoTerima = IIf(IsNull(rsb("NoFIFO")), "0000000000", rsb("NoFIFO"))
                                    'dblJmlStokMax = dbRst("JmlStokMax")
                                    strNoTerima = rsC("NoFIFO")
                                    .TextMatrix(intRowTemp, 31) = strNoTerima

                                    strSQL = ""
                                    Set rsB = Nothing
                                    strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 6) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 31) & "') AS HargaBarang"
                                    Call msubRecFO(rsB, strSQL)
                                    If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

                                    strSQL = ""
                                    Set rsB = Nothing
                                    subcurHargaSatuan = 0

                                    strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                    Call msubRecFO(rsB, strSQL)
                                    If rsB.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rsB(0).Value
                                    'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                                    'subcurHargaSatuan = (subcurHargaSatuan * typSettingDataPendukung.realPPn / 100) + subcurHargaSatuan
                                    .TextMatrix(intRowTemp, 7) = subcurHargaSatuan
                                    .TextMatrix(intRowTemp, 7) = Format(subcurHargaSatuan, "#,###")
                                    .TextMatrix(intRowTemp, 8) = (.TextMatrix(.Row, 8) / 100) * subcurHargaSatuan

                                    .TextMatrix(intRowTemp, 10) = Abs(dblSelisih)

                                    dblSelisih = Abs(dblSelisih) - CDbl(rsC("JmlStokMax"))
                                    If dblSelisih >= 0 Then
                                        '.TextMatrix(intRowTemp, 9) = rsc("JmlStokMax")
                                        .TextMatrix(intRowTemp, 10) = rsC("JmlStokMax")
                                    End If

                                    .TextMatrix(intRowTemp, 13) = .TextMatrix(.Row, 13)

                                    .TextMatrix(intRowTemp, 23) = ""
                                    .TextMatrix(intRowTemp, 24) = curHargaBrg
                                    .TextMatrix(intRowTemp, 25) = .TextMatrix(.Row, 25)
                                    .TextMatrix(intRowTemp, 26) = 0
                                    .TextMatrix(intRowTemp, 27) = .TextMatrix(.Row, 27)
                                    .TextMatrix(intRowTemp, 28) = .TextMatrix(.Row, 28)

                                    .TextMatrix(intRowTemp, 14) = .TextMatrix(.Row, 14)
                                    .TextMatrix(intRowTemp, 15) = .TextMatrix(.Row, 15)
                                    .TextMatrix(intRowTemp, 16) = .TextMatrix(intRowTemp, 7)

                                    'ambil no temporary
                                    If sp_TempDetailApotikJual(CDbl(.TextMatrix(intRowTemp, 7)) + CDbl(.TextMatrix(intRowTemp, 14)), .TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12)) = False Then Exit Sub
                                    'ambil hutang penjamin dan tanggungan rs
                                    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
                                    " FROM TempDetailApotikJual" & _
                                    " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(intRowTemp, 2) & "') AND (KdAsal = '" & .TextMatrix(intRowTemp, 12) & "')"
                                    Call msubRecFO(rsB, strSQL)
                                    If rsB.EOF = True Then
                                        curHutangPenjamin = 0
                                        curTanggunganRS = 0
                                    Else
                                        curHutangPenjamin = rsB("JmlHutangPenjamin").Value
                                        curTanggunganRS = rsB("JmlTanggunganRS").Value
                                    End If

                                    .TextMatrix(intRowTemp, 11) = ((CDbl(.TextMatrix(intRowTemp, 14)) * CDbl(.TextMatrix(intRowTemp, 15))) + _
                                    (CDbl(.TextMatrix(intRowTemp, 16)) * Val(.TextMatrix(intRowTemp, 10)))) + Val(.TextMatrix(intRowTemp, 26))

                                    .TextMatrix(intRowTemp, 17) = curHutangPenjamin
                                    .TextMatrix(intRowTemp, 18) = curTanggunganRS
                                    If .TextMatrix(intRowTemp, 17) > 0 Then
                                        .TextMatrix(intRowTemp, 19) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 17))) + Val(.TextMatrix(intRowTemp, 26))
                                    Else
                                        .TextMatrix(intRowTemp, 19) = 0
                                    End If

                                    If .TextMatrix(intRowTemp, 18) > 0 Then
                                        .TextMatrix(intRowTemp, 20) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 18))) + Val(.TextMatrix(intRowTemp, 26))
                                    Else
                                        .TextMatrix(intRowTemp, 20) = 0
                                    End If
                                    .TextMatrix(intRowTemp, 21) = ((CDbl(.TextMatrix(intRowTemp, 8) / 100)) * (CDbl(.TextMatrix(intRowTemp, 10)))) '* CDbl(.TextMatrix(intRowTemp, 16))))

                                    curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + _
                                    CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                    .TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                    If dblSelisih <= 0 Then Exit For

                                    rsC.MoveNext
                                Next k
                            End If
                        End With
                    Else
                        bolCekFIFO = False
                    End If
                End If

                If bolStatusFIFO = True Then i = iTemp
                .Rows = .Rows + 1
                dbRst.MoveNext
            Next i
            Call subHitungTotal
        End With
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errLoad
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    frmTransaksiPasien.Enabled = True
errLoad:
End Sub

Private Sub subHapusDataGrid()
    On Error GoTo errLoad
    Dim i As Integer
    Dim strResepKe As String
    Dim intBarisYangDihapus As Integer
    Dim curHarusDibayar As Currency

    With fgData
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, 11)) = 0 Then GoTo stepHapusData
        intBarisYangDihapus = fgData.Row
        If .TextMatrix(.Row, 11) <> "01" Then 'jika obat racikan, pastikan jumlah service 1 untuk resep yang sama
            strResepKe = .TextMatrix(.Row, 0)
            If Val(.TextMatrix(.Row, 15)) = 0 Then GoTo stepHapusData
            For i = 1 To .Rows - 2
                If .TextMatrix(i, 0) = strResepKe And i <> intBarisYangDihapus Then
                    .TextMatrix(i, 13) = 1
                    Exit For
                End If
            Next i
        End If

stepHapusData:
        'add by onede
        ' If sp_StokRuangan(.TextMatrix(.Row, 2), .TextMatrix(.Row, 12), CDbl(.TextMatrix(.Row, 10)), "A") = False Then Exit Sub

        dbConn.Execute "DELETE FROM TempDetailApotikJual " & _
        " WHERE (NoTemporary = '" & Trim(.TextMatrix(.Row, 23)) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
        If .Rows = 2 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .TextMatrix(1, 0) = "1"
            Next i
        Else
            .RemoveItem .Row
        End If
        'If .TextMatrix(.Row, 10) = "" Then .TextMatrix(.Row, 10) = 0
        If .TextMatrix(.Row, 2) <> "" Then
            .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
            (CDbl(.TextMatrix(.Row, 16)) * CDbl(.TextMatrix(.Row, 10))))

            curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
            CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
            .TextMatrix(.Row, 20) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
        End If
    End With
    Call subHitungTotal

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Private Sub Timer1_Timer()
'dtpTglPelayanan.Value = Now
'End Sub

Private Sub txtDokter_Change()
    On Error GoTo errLoad
    mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    txtKdDokter.Text = ""
    dgDokter.Visible = True
    Call subLoadDokter
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtDokter_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, vbKeyDown
            If dgDokter.Visible = True Then dgDokter.SetFocus Else fgData.SetFocus
        Case vbKeyEscape
            dgDokter.Visible = False
    End Select
End Sub


Private Sub txtIsi_Change()
    Dim i As Integer
    Select Case fgData.Col
        Case 2 'kode barang
            If tempStatusTampil = True Then Exit Sub
            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
            Call msubRecFO(dbRst, strSQL)

            Set dgObatAlkes.DataSource = dbRst
            With dgObatAlkes
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i

                .Columns("KdBarang").Width = 1500
                .Columns("NamaBarang").Width = 3000
                .Columns("JenisBarang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675

                .Top = 2830
                .Left = 1820
                .Visible = True
                For i = 1 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
                If fgData.TopRow > 1 Then
                    .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                End If
            End With
        Case 3 ' nama barang
            If tempStatusTampil = True Then Exit Sub
            strSQL = "execute CariBarangNStokMedis_V '" & txtIsi.Text & "%','" & mstrKdRuangan & "'"
            Call msubRecFO(dbRst, strSQL)

            Set dgObatAlkes.DataSource = dbRst
            With dgObatAlkes
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next i

                .Columns("KdBarang").Width = 1500
                .Columns("NamaBarang").Width = 3000
                .Columns("JenisBarang").Width = 1500
                .Columns("Kekuatan").Width = 1000
                .Columns("AsalBarang").Width = 1000
                .Columns("Satuan").Width = 675

                .Top = 2830
                .Left = 3000
                .Visible = True
                For i = 1 To fgData.Row - 1
                    .Top = .Top + fgData.RowHeight(i)
                Next i
                If fgData.TopRow > 1 Then
                    .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                End If
            End With

        Case Else
            dgObatAlkes.Visible = False
    End Select
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then If dgObatAlkes.ApproxCount > 0 Then dgObatAlkes.SetFocus

End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i, j As Integer
    Dim curHutangPenjamin As Currency
    Dim curTanggunganRS As Currency
    Dim curHarusDibayar As Currency
    Dim KdJnsObat As String
    Dim dblSelisih As Double
    Dim intRowTemp As Integer
    Dim strNoTerima As String
    Dim curHargaBrg As Currency
    Dim dblSelisihNow As Double
    Dim dblJmlStokMax As Double
    Dim strKdBrg As String
    Dim strKdAsal As String
    Dim dblJmlTerkecil As Double
    Dim dblTotalStokK As Double
        With fgData
            Select Case fgData.Col
                Case 0
                    Call SetKeyPressToNumber(KeyAscii)
                Case 10
                    Call SetKeyPressToNumber(KeyAscii)
            End Select
        End With

    If KeyAscii = 13 Then
    
        With fgData
            Select Case fgData.Col
                Case 0
                    Call SetKeyPressToNumber(KeyAscii)
                Case 10
                    
                    If fgData.TextMatrix(fgData.Row, 1) = "" Then
                        MsgBox "Jenis Obat harus diisi", vbExclamation, "Validasi": Exit Sub
                    End If
                        
                    If fgData.TextMatrix(fgData.Row, 3) = "" Then
                        MsgBox "Nama Obat harus diisi", vbExclamation, "Validasi": Exit Sub
                    End If
                    
                    Call SetKeyPressToNumber(KeyAscii)
            End Select
        End With
    
        With fgData
            Select Case .Col
                Case 0
                    dgObatAlkes.Visible = False
                    If Val(txtIsi.Text) = 0 Then txtIsi.Text = 1
                    .TextMatrix(.Row, .Col) = CDbl(txtIsi.Text)
                    txtIsi.Visible = False

                    dcJenisObat.Left = 120
                    .Col = 1
                    For i = 0 To .Col - 1
                        dcJenisObat.Left = dcJenisObat.Left + .ColWidth(i)
                    Next i
                    dcJenisObat.Visible = True
                    dcJenisObat.Top = .Top - 7

                    For i = 0 To .Row - 1
                        dcJenisObat.Top = dcJenisObat.Top + .RowHeight(i)
                    Next i

                    If .TopRow > 1 Then
                        dcJenisObat.Top = dcJenisObat.Top - ((.TopRow - 1) * .RowHeight(1))
                    End If

                    dcJenisObat.Width = .ColWidth(.Col)
                    dcJenisObat.Height = .RowHeight(.Row)

                    dcJenisObat.Visible = True
                    dcJenisObat.SetFocus

                Case 1

                Case 2
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        fgData.SetFocus
                        fgData.Col = 8
                    End If

                Case 3
                    If dgObatAlkes.Visible = True Then
                        dgObatAlkes.SetFocus
                        Exit Sub
                    Else
                        fgData.SetFocus
                        fgData.Col = 8
                    End If

                Case 8
                    dgObatAlkes.Visible = False
                    txtIsi.Visible = False

                    If mblnOperator = False Then
                        If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0

                        'konvert koma col discount
                        .TextMatrix(.Row, .Col) = Val(txtIsi.Text)
                        .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col)

                        'konvert koma col jumlah
                        .TextMatrix(.Row, 10) = CDbl(.TextMatrix(.Row, 10))

                        If .TextMatrix(.Row, 10) <> "0" Then
                            .TextMatrix(.Row, 21) = IIf(CDbl(CDbl(.TextMatrix(.Row, 10))) = 0, 0, CDbl(.TextMatrix(.Row, 10))) * CDbl(.TextMatrix(.Row, 8))

                            curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
                            (CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20))))
                            .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)
                            Call subHitungTotal
                        End If

                    End If

                    fgData.SetFocus
                    fgData.Col = 10

                Case 10
                    If Trim(txtIsi.Text) = "," Then txtIsi.Text = 0
                    If Trim(txtIsi.Text) = "" Then txtIsi.Text = 0
                    If (fgData.TextMatrix(.Row, 6) = "S") Then
                       If bolStatusFIFO = False Then
                        If CDbl(txtIsi.Text) > CDbl(.TextMatrix(.Row, 9)) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                       End If
                    ElseIf (fgData.TextMatrix(.Row, 6) = "K") Then
                        Set rs = Nothing
                        strSQL = "Select JmlTerkecil From MasterBarang Where KdBarang = '" & fgData.TextMatrix(.Row, 2) & "'"
                        Call msubRecFO(rs, strSQL)
                        dblJmlTerkecil = IIf(rs.EOF, 1, rs(0).Value)

                        dblTotalStokK = dblJmlTerkecil * fgData.TextMatrix(.Row, 9)
                        If Val(txtIsi.Text) > Val(dblTotalStokK) Then
                            MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 9) & ")", vbExclamation, "Validasi"
                            txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
                            Exit Sub
                        End If
                    End If

                    'add for FIFO validasi jika terjadi edit jml stok, hapus otomatis
                    If bolStatusFIFO = True Then
                        If Trim(.TextMatrix(.Row, 10)) <> "" Then
                            i = .Rows - 1
                            strKdBrg = .TextMatrix(.Row, 2)
                            strKdAsal = .TextMatrix(.Row, 12)
                            Do While i <> 1
                                If .TextMatrix(i, 2) <> "" Then
                                    If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                        .Row = i
                                        If .CellBackColor = vbRed Then
                                            Call subHapusDataGrid
                                            .Row = i - 1
                                        End If
                                    End If
                                End If
                                i = i - 1
                            Loop

                            For i = 1 To .Rows - 1
                                If (strKdBrg = .TextMatrix(i, 2)) And (strKdAsal = .TextMatrix(i, 12)) Then
                                    .Row = i
                                    Exit For
                                End If
                            Next i
                        End If

                        .SetFocus
                        intRowTemp = 0
                    End If

                    'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                    If bolStatusFIFO = True Then
                        Set dbRst = Nothing
                      '  Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "','" & .TextMatrix(.Row, 31) & "') as stok")
                        Call msubRecFO(dbRst, "select dbo.FB_TakeStokBrgMedis('" & mstrKdRuangan & "', '" & .TextMatrix(.Row, 2) & "','" & .TextMatrix(.Row, 12) & "') as stok")
                        If .TextMatrix(.Row, 6) = "S" Then
                            dblSelisih = dbRst(0) - CDbl(txtIsi.Text)
                        Else
                            dblSelisih = (dbRst(0) * dblJmlTerkecil) - CDbl(txtIsi.Text)
                        End If
                        If dblSelisih < 0 Then
                            If .TextMatrix(.Row, 6) = "S" Then
                                txtIsi.Text = dbRst(0)
                            Else
                                txtIsi.Text = dbRst(0) * dblJmlTerkecil
                            End If
                            '.TextMatrix(.Row, 9) = dbRst(0)
                        Else
'                            Set dbRst = Nothing
'                            If bolStatusFIFO = False Then
'                                strSQL = "Select JmlStok as Stok From StokRuangan Where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and KdRuangan='" & mstrKdRuangan & "'"
'                               Else
'                                strSQL = "Select JmlStok as Stok From StokRuanganFIFO Where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and KdRuangan='" & mstrKdRuangan & "' and Noterima ='" & .TextMatrix(.Row, 31) & "'"
'                            End If
'                            Call msubRecFO(dbRst, strSQL)
'                            If dbRst.EOF Then
'                                .TextMatrix(.Row, 9) = 0
'                            Else
'                                .TextMatrix(.Row, 9) = IIf(IsNull(dbRst("Stok")), 0, dbRst("Stok"))
'                            End If
                        End If
                    End If
                    'end FIFO

                    .TextMatrix(.Row, .Col) = txtIsi.Text

                    'konvert koma col discount
                    .TextMatrix(.Row, 8) = .TextMatrix(.Row, 8)

                    txtIsi.Visible = False

                    '                    If dcJenisObat.Text = "01" Then
                    '                        subintJmlService = 1 'CDbl(.TextMatrix(.Row, 8))
                    '                    Else
                    subintJmlService = 1 'default baris pertama
                    '                    End If
                    'rubah jumlah service
                    .TextMatrix(.Row, 15) = subintJmlService

                    'ambil no temporary
                    If sp_TempDetailApotikJual(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 14)) + Val(.TextMatrix(i, 26)), .TextMatrix(.Row, 2), .TextMatrix(.Row, 12)) = False Then Exit Sub
                    'ambil hutang penjamin dan tanggungan rs
                    strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
                    " FROM TempDetailApotikJual" & _
                    " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(.Row, 2) & "') AND (KdAsal = '" & .TextMatrix(.Row, 12) & "')"
                    Call msubRecFO(rs, strSQL)
                    If rs.EOF = True Then
                        curHutangPenjamin = 0
                        curTanggunganRS = 0
                    Else
                        curHutangPenjamin = rs("JmlHutangPenjamin").Value
                        curTanggunganRS = rs("JmlTanggunganRS").Value
                    End If

                    'rubah jumlah service
                    If dcJenisObat.BoundText = "01" Then .TextMatrix(.Row, 15) = 1 'val(.TextMatrix(.Row, 10))

                    .TextMatrix(.Row, 14) = subcurTarifService
                    .TextMatrix(.Row, 16) = CDbl(.TextMatrix(.Row, 7))

                    'total harga = ((tarifservice * jmlservice) + _
                    (hargasatuan(sebelum ditambah tarifservixe) * jumlah))
                    .TextMatrix(.Row, 11) = ((CDbl(.TextMatrix(.Row, 14)) * CDbl(.TextMatrix(.Row, 15))) + _
                    (CCur(.TextMatrix(.Row, 16)) * CDbl(.TextMatrix(.Row, 10)))) + Val(.TextMatrix(.Row, 26))
                    '                    .Col = 11: .CellForeColor = vbBlue: .CellFontBold = True: .Col = 10

                    .TextMatrix(.Row, 17) = curHutangPenjamin
                    .TextMatrix(.Row, 18) = curTanggunganRS

                    If curHutangPenjamin > 0 Then
                        .TextMatrix(.Row, 19) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 17))) + Val(.TextMatrix(.Row, 26))
                    Else
                        .TextMatrix(.Row, 19) = 0
                    End If

                    If curTanggunganRS > 0 Then
                        .TextMatrix(.Row, 20) = (.TextMatrix(.Row, 14) * .TextMatrix(.Row, 15)) + (CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 18))) + Val(.TextMatrix(i, 26))
                    Else
                        .TextMatrix(.Row, 20) = 0
                    End If
                    .TextMatrix(.Row, 21) = CDbl(.TextMatrix(.Row, 10)) * CDbl(.TextMatrix(.Row, 8))

                    'total harus dibayar = total harga - total discount - _
                    total hutang penjamin - totaltanggunganrs
                    curHarusDibayar = CDbl(.TextMatrix(.Row, 11)) - (CDbl(.TextMatrix(.Row, 21)) + _
                    CDbl(.TextMatrix(.Row, 19)) + CDbl(.TextMatrix(.Row, 20)))
                    .TextMatrix(.Row, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                    'add for FIFO jika jml yg diinput melebihi stok penerimaan, mk otomatis muncul di row selanjutnya
                    
                    If bolStatusFIFO = True Then
                        If dblSelisih < 0 Then
                            With fgData
'                                strSQL = "select NoTerima As NoFIFO,JmlStokMax from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 31) & "' and JmlStok<>0 order by TglTerima asc"
                              strSQL = "select NoTerima As NoFIFO,JmlStokMax,JmlStok from V_StokRuanganFIFO where KdBarang='" & .TextMatrix(.Row, 2) & "' and KdAsal='" & .TextMatrix(.Row, 12) & "' and NoTerima<>'" & .TextMatrix(.Row, 31) & "' and JmlStok<>0 and Kdruangan='" & mstrKdRuangan & "' order by TglTerima asc"
 
                                Set dbRst = Nothing
                                Call msubRecFO(dbRst, strSQL)
                                If dbRst.EOF = False Then
                                    dbRst.MoveFirst
                                    For i = 1 To dbRst.RecordCount
                                        '.Rows = .Rows - 1

                                        .Rows = .Rows + 1

                                        intRowTemp = .Row
                                        If .TextMatrix(.Rows - 2, 2) = "" Then
                                            .Row = .Rows - 2
                                        Else
                                            .Row = .Rows - 1
                                        End If
                                        For j = 0 To .Cols - 1
                                            .Col = j
                                            .CellBackColor = vbRed
                                            .CellForeColor = vbWhite
                                        Next j

                                        .Row = intRowTemp
                                        intRowTemp = 0
                                        If .TextMatrix(.Rows - 2, 2) = "" Then
                                            intRowTemp = .Rows - 2
                                        Else
                                            intRowTemp = .Rows - 1
                                        End If

                                        curHutangPenjamin = 0
                                        curTanggunganRS = 0

                                        .TextMatrix(intRowTemp, 0) = .TextMatrix(.Row, 0)
                                        .TextMatrix(intRowTemp, 1) = .TextMatrix(.Row, 1)
                                        .TextMatrix(intRowTemp, 2) = .TextMatrix(.Row, 2)
                                        .TextMatrix(intRowTemp, 3) = .TextMatrix(.Row, 3)
                                        .TextMatrix(intRowTemp, 4) = .TextMatrix(.Row, 4)
                                        .TextMatrix(intRowTemp, 5) = .TextMatrix(.Row, 5)
                                        .TextMatrix(intRowTemp, 6) = .TextMatrix(.Row, 6)
                                        .TextMatrix(intRowTemp, 12) = .TextMatrix(.Row, 12)

                                        strNoTerima = dbRst("NoFIFO")
                                        .TextMatrix(intRowTemp, 31) = strNoTerima

                                        strSQL = ""
                                        Set rsB = Nothing
                                        strSQL = "SELECT dbo.FB_TakeHargaNettoOA('" & mstrKdPenjaminPasien & "','" & mstrKdJenisPasien & "','" & .TextMatrix(intRowTemp, 2) & "','" & .TextMatrix(intRowTemp, 12) & "','" & .TextMatrix(intRowTemp, 6) & "', '" & mstrKdRuangan & "','" & .TextMatrix(intRowTemp, 31) & "') AS HargaBarang"
                                        Call msubRecFO(rsB, strSQL)
                                        If rsB.EOF = True Then curHargaBrg = 0 Else curHargaBrg = rsB(0).Value

                                        strSQL = ""
                                        Set rs = Nothing
                                        subcurHargaSatuan = 0
                            
                                        strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & .TextMatrix(intRowTemp, 12) & "', " & msubKonversiKomaTitik(CStr(curHargaBrg)) & ")  as HargaSatuan"
                                        Call msubRecFO(rs, strSQL)
                                        If rs.EOF = True Then subcurHargaSatuan = 0 Else subcurHargaSatuan = rs(0).Value
                                        'khusus OA harga tidak dikalikan lg Ppn krn OA termasuk pelayanan yg include ke tindakan (TM)
                                        'subcurHargaSatuan = (subcurHargaSatuan * typSettingDataPendukung.realPPn / 100) + subcurHargaSatuan
                                        .TextMatrix(intRowTemp, 7) = subcurHargaSatuan
                                        .TextMatrix(intRowTemp, 7) = Format(subcurHargaSatuan, "#,###")
                                        .TextMatrix(intRowTemp, 8) = (.TextMatrix(.Row, 8) / 100) * subcurHargaSatuan

                                        .TextMatrix(intRowTemp, 10) = Abs(dblSelisih)

                                        If .TextMatrix(intRowTemp, 6) = "S" Then
                                            dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStokMax"))
                                        Else
                                            dblSelisih = Abs(dblSelisih) - CDbl(dbRst("JmlStokMax") * dblJmlTerkecil)
                                        End If
                                        If dblSelisih >= 0 Then
                                            '.TextMatrix(intRowTemp, 9) = dbRst("JmlStokMax")
                                            If .TextMatrix(intRowTemp, 6) = "S" Then
                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStokMax")
                                            Else
                                                .TextMatrix(intRowTemp, 10) = dbRst("JmlStokMax") * dblJmlTerkecil
                                            End If
                                        End If
                                        .TextMatrix(intRowTemp, 9) = dbRst("JmlStok")
                                        .TextMatrix(intRowTemp, 13) = .TextMatrix(.Row, 13)

                                        .TextMatrix(intRowTemp, 23) = ""
                                        .TextMatrix(intRowTemp, 24) = curHargaBrg
                                        .TextMatrix(intRowTemp, 25) = .TextMatrix(.Row, 25)
                                        .TextMatrix(intRowTemp, 26) = 0
                                        .TextMatrix(intRowTemp, 27) = .TextMatrix(.Row, 27)
                                        .TextMatrix(intRowTemp, 28) = .TextMatrix(.Row, 28)
                                        .TextMatrix(intRowTemp, 29) = .TextMatrix(.Row, 29)

                                        .TextMatrix(intRowTemp, 14) = .TextMatrix(.Row, 14)
                                        .TextMatrix(intRowTemp, 15) = .TextMatrix(.Row, 15)
                                        .TextMatrix(intRowTemp, 16) = .TextMatrix(intRowTemp, 7)

                                        'ambil no temporary
                                        If sp_TempDetailApotikJual(CDbl(.TextMatrix(intRowTemp, 7)) + CDbl(.TextMatrix(intRowTemp, 14)), .TextMatrix(intRowTemp, 2), .TextMatrix(intRowTemp, 12)) = False Then Exit Sub
                                        'ambil hutang penjamin dan tanggungan rs
                                        strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
                                        " FROM TempDetailApotikJual" & _
                                        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & .TextMatrix(intRowTemp, 2) & "') AND (KdAsal = '" & .TextMatrix(intRowTemp, 12) & "')"
                                        Call msubRecFO(rs, strSQL)
                                        If rs.EOF = True Then
                                            curHutangPenjamin = 0
                                            curTanggunganRS = 0
                                        Else
                                            curHutangPenjamin = rs("JmlHutangPenjamin").Value
                                            curTanggunganRS = rs("JmlTanggunganRS").Value
                                        End If

                                        .TextMatrix(intRowTemp, 11) = ((CDbl(.TextMatrix(intRowTemp, 14)) * CDbl(.TextMatrix(intRowTemp, 15))) + _
                                        (CDbl(.TextMatrix(intRowTemp, 16)) * Val(.TextMatrix(intRowTemp, 10)))) + Val(.TextMatrix(intRowTemp, 26))

                                        .TextMatrix(intRowTemp, 17) = curHutangPenjamin
                                        .TextMatrix(intRowTemp, 18) = curTanggunganRS
                                        If .TextMatrix(intRowTemp, 17) > 0 Then
                                            .TextMatrix(intRowTemp, 19) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 17))) + Val(.TextMatrix(intRowTemp, 26))
                                        Else
                                            .TextMatrix(intRowTemp, 19) = 0
                                        End If

                                        If .TextMatrix(intRowTemp, 18) > 0 Then
                                            .TextMatrix(intRowTemp, 20) = (.TextMatrix(intRowTemp, 14) * .TextMatrix(intRowTemp, 15)) + (Val(.TextMatrix(intRowTemp, 10)) * CDbl(.TextMatrix(intRowTemp, 18))) + Val(.TextMatrix(intRowTemp, 26))
                                        Else
                                            .TextMatrix(intRowTemp, 20) = 0
                                        End If
                                        .TextMatrix(intRowTemp, 21) = ((CDbl(.TextMatrix(intRowTemp, 8) / 100)) * (CDbl(.TextMatrix(intRowTemp, 10)))) '* CDbl(.TextMatrix(intRowTemp, 16))))

                                        curHarusDibayar = CDbl(.TextMatrix(intRowTemp, 11)) - (CDbl(.TextMatrix(intRowTemp, 21)) + _
                                        CDbl(.TextMatrix(intRowTemp, 19)) + CDbl(.TextMatrix(intRowTemp, 20)))
                                        .TextMatrix(intRowTemp, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

                                        If dblSelisih <= 0 Then Exit For

                                        dbRst.MoveNext
                                    Next i
                                End If
                            End With
                        End If
                    End If
                    'end fifo

                    Call subHitungTotal
                    fgData.SetFocus
                    fgData.Col = 28
                    ' Call subLoadCheck
            End Select
        End With

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    ElseIf (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Or KeyAscii = vbKeySpace) Then
        If fgData.Col <> 3 Then
            dgObatAlkes.Visible = False
        Else
            dgObatAlkes.Visible = True
            txtIsi.Visible = True
        End If
    Else
        KeyAscii = 0
    End If

    If KeyAscii < 48 Or KeyAscii > 58 Then
        With fgData
            If KeyAscii = 8 Then
                Select Case .Col
                    Case 10
                        KeyAscii = 8
                End Select
            Else
                Select Case .Col
                    Case 10
                        KeyAscii = 0
                End Select
            End If
        End With
    End If

    Exit Sub
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad
    With fgData
        .clear
        .Rows = 2
        .Cols = 32

        .RowHeight(0) = 400

        .TextMatrix(0, 0) = "R/Ke"
        .TextMatrix(1, 0) = "1"
        .TextMatrix(0, 1) = "Jenis Obat"
        .TextMatrix(0, 2) = "KodeBarang"
        .TextMatrix(0, 3) = "Nama Barang"
        .TextMatrix(0, 4) = "Kekuatan"
        .TextMatrix(0, 5) = "Asal Barang"
        .TextMatrix(0, 6) = "Satuan"
        .TextMatrix(0, 7) = "Harga Satuan" 'udah ditambah tarif service
        .TextMatrix(0, 8) = "Discount" 'per 1 barang
        .TextMatrix(0, 9) = "Stock" 'untuk perbandingan jika jumlah dirubah cek stok
        .TextMatrix(0, 10) = "Jumlah"
        .TextMatrix(0, 11) = "Total Harga" 'total = ((tarifservice * jmlservice) + (hargasatuan(sebelum ditambah tarifservixe) * jumlah))

        .TextMatrix(0, 12) = "KdAsal"
        .TextMatrix(0, 13) = "Jenis Barang"

        .TextMatrix(0, 14) = "TarifServise" 'per jenis obat
        .TextMatrix(0, 15) = "JmlService" 'jika obat jadi = Jumlah, 1 atau 0
        .TextMatrix(0, 16) = "HargaSebelumTarifService" 'harga satuan sesudah take tarif dan sebelum ditambah tarif service

        .TextMatrix(0, 17) = "HutangPenjamin" 'PER BARANG, diambil dari TempDetailApotikJual
        .TextMatrix(0, 18) = "TanggunganRS" 'PER BARANG, diambil dari TempDetailApotikJual

        .TextMatrix(0, 19) = "TotalHutangPenjamin" 'jumlah * hutang penjamin
        .TextMatrix(0, 20) = "TotalTanggunganRS" 'jumlah * TotalTanggunganRS
        .TextMatrix(0, 21) = "TotalDiscount" 'jumlah * discount 1 barang
        .TextMatrix(0, 22) = "TotalHarusBayar" 'curHarusDibayar = Total Harga - TotalDiscount - TotalHutangPenjamin - TotalTanggunganRS _
        iif curHarusDibayar < 0, 0, curHarusDibayar
        .TextMatrix(0, 23) = "NoTemp" 'jika barang dihapus digrid, hapus ke tabel TempDetailApotikJual
        .TextMatrix(0, 24) = "HargaBeli" 'harga satuan sebelum take tarif dan sebelum ditambah tarif service
        .TextMatrix(0, 25) = "KdJenisObat"
        .TextMatrix(0, 26) = "BiayaAdministrasi"

        .TextMatrix(0, 27) = "Kirim"

        .TextMatrix(0, 28) = "Pemakaian Pemeriksaan"
        .TextMatrix(0, 29) = "KodePelayananRS"
        .TextMatrix(0, 30) = "StatusOrder"
        .TextMatrix(0, 31) = "NoTerima"

        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 3200
        .ColWidth(4) = 0
        .ColWidth(5) = 1100
        .ColWidth(6) = 0
        .ColWidth(7) = 1200
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 1200
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 0
        .ColWidth(23) = 0
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0
        .ColWidth(28) = 2000
        .ColWidth(29) = 0
        .ColWidth(30) = 0
        .ColWidth(31) = 0

        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    'Call msubDcSource(dcJenisObat, rs, "SELECT KdJenisObat, JenisObat FROM JenisObat where StatusEnabled='1' ORDER BY JenisObat")
    Call msubDcSource(dcJenisObat, rs, "SELECT KdJenisObat, JenisObat FROM JenisObat where KdJenisObat in ('02','03') ORDER BY JenisObat")
    If rs.EOF = False Then dcJenisObat.BoundText = rs(0).Value

    strSQL = "SELECT  TOP (200) KdPelayananRS, NamaPelayanan, NoPendaftaran" & _
    " FROM  V_NamaPelayananPerPasien where NoPendaftaran ='" & mstrNoPen & "'"
    Call msubDcSource(dcNamaPelayananRS, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJmlTerima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpanTerimaBarang.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtNoResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglResep.SetFocus
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
    On Error GoTo errLoad

    strSQL = "SELECT NamaDokter AS [Nama Dokter],JK,Jabatan,KodeDokter  FROM V_DaftarDokter " & mstrFilterDokter
    Call msubRecFO(rs, strSQL)
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 3500
        .Columns(1).Width = 400
        .Columns(2).Width = 1600
        .Columns(3).Width = 0
    End With
    dgDokter.Left = 5760
    dgDokter.Top = 1920

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function f_HitungTotal() As Currency
    On Error GoTo errLoad
    Dim i As Integer

    f_HitungTotal = 0
    For i = 1 To fgData.Rows - 2
        f_HitungTotal = f_HitungTotal + fgData.TextMatrix(i, 11)
    Next i

    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function sp_ResepObat() As Boolean
    On Error GoTo errLoad

    sp_ResepObat = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, txtNoResep.Text)
        .Parameters.Append .CreateParameter("TglResep", adDate, adParamInput, , Format(dtpTglResep.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(txtKdDokter.Text = "", Null, txtKdDokter.Text))
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(mstrKdRuanganPasien = "", Null, mstrKdRuanganPasien))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("ResepBebas", adChar, adParamInput, 1, "T")

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_ResepObat"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_ResepObat = False

        End If
    End With

    Exit Function
errLoad:
    Call msubPesanError
    sp_ResepObat = False
End Function

Public Function sp_PenerimaanSementara(f_Tanggal As Date, f_KdBarang As String, f_KdAsal As String, f_JmlBarang As Double, f_status As String) As Boolean
    On Error GoTo errLoad
    Dim i As Integer

    sp_PenerimaanSementara = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglTerima", adDate, adParamInput, , Format(f_Tanggal, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("JmlTerima", adInteger, adParamInput, , f_JmlBarang)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.add_PenerimaanBarangApotikTemp"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "validasi"
            sp_PenerimaanSementara = False
        Else
            Call Add_HistoryLoginActivity("add_PenerimaanBarangApotikTemp")
        End If
    End With
    Set dbcmd = Nothing

    Exit Function
errLoad:
    sp_PenerimaanSementara = False
    Call msubPesanError("sp_PenerimaanSementara")
End Function

Private Function sp_PemakaianObatAlkesResep(f_KdBarang As String, f_KdAsal As String, _
    f_Satuan As String, f_Jumlah As Double, f_HargaSebelumTarifService As Currency, f_KdJenisObat As String, _
    f_JumlahServise As Integer, f_TarifService As Currency, f_Rke As Integer, f_StatusStok As String, _
    f_KdPelayananUsed As String, f_KdStatusHasil As String, f_JmlExpose As String, f_KdStatusKontras As String, f_idPenanggungjawab As String, f_Keterangan As String, f_NoTerima As String, f_tglPelayanan As Date) As Boolean
    On Error GoTo errLoad
    Dim i As Integer
    sp_PemakaianObatAlkesResep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlBrg", adDouble, adParamInput, , CDbl(f_Jumlah))
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSebelumTarifService)
        '.Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(mstrKdDokter = "", strIDPegawaiAktif, mstrKdDokter))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, IIf(f_KdJenisObat = "", Null, f_KdJenisObat))
        .Parameters.Append .CreateParameter("JmlService", adInteger, adParamInput, , f_JumlahServise)
        .Parameters.Append .CreateParameter("TarifService", adCurrency, adParamInput, , f_TarifService)
        .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, IIf(chkNoResep.Value = vbChecked, txtNoResep.Text, Null))
        .Parameters.Append .CreateParameter("Rke", adInteger, adParamInput, , IIf(Len(Trim(f_Rke)) = 0, Null, f_Rke))
        .Parameters.Append .CreateParameter("StatusStok", adChar, adParamInput, 1, f_StatusStok)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, StrKdRP)

        .Parameters.Append .CreateParameter("KdPelayananRSUsed", adChar, adParamInput, 6, IIf(f_KdPelayananUsed = "", Null, f_KdPelayananUsed))
        .Parameters.Append .CreateParameter("KdStatusHasil", adChar, adParamInput, 2, IIf(f_KdStatusHasil = "", Null, f_KdStatusHasil))
        .Parameters.Append .CreateParameter("JmlExpose", adInteger, adParamInput, , IIf(f_JmlExpose = "", Null, f_JmlExpose))
        .Parameters.Append .CreateParameter("KdStatusKontras", adInteger, adParamInput, , IIf(f_KdStatusKontras = "", Null, f_KdStatusKontras))
        .Parameters.Append .CreateParameter("IdPenanggungjawab", adChar, adParamInput, 1, IIf(f_idPenanggungjawab = "", Null, f_idPenanggungjawab))
        .Parameters.Append .CreateParameter("Keterangan", adChar, adParamInput, 3, IIf(f_Keterangan = "", Null, f_Keterangan))
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PemakaianObatAlkesResepNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PemakaianObatAlkesResep = False

        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    ' add by onede
    'untuk simpan ulang jika error(Time expired)
    Call deleteADOCommandParameters(dbcmd)
    For i = 1 To fgData.Rows - 2
        If fgData.TextMatrix(i, 2) <> "" Then
            strSQL = "SELECT  NoPendaftaran, KdRuangan, KdBarang, KdAsal, TglPelayanan FROM  PemakaianAlkes" & _
            " WHERE NoPendaftaran = '" & mstrNoPen & "'  AND KdRuangan ='" & mstrKdRuangan & "' AND KdBarang ='" & fgData.TextMatrix(i, 2) & "'  AND KdAsal ='" & fgData.TextMatrix(i, 12) & "'" & _
            "and day(TglPelayanan)=day('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(dtpTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "')"
            Call msubRecFO(rs, strSQL)

            If rs.EOF = False Then fgData.RemoveItem i
        End If
    Next i

    MsgBox "Waktu Penyimpanan Habis..Tekan kembali tombol simpan untuk menyimpan barang yg belum tersimpan!!!", vbExclamation, "Validasi"
    sp_PemakaianObatAlkesResep = False
End Function

Private Function sp_EtiketResep(f_KdBarang As String, f_KdAsal As String, f_KdJenisObat As String, _
    f_Signa As String, f_KdSatuanEtiket As String, f_KdWaktuEtiket As String, f_ResepKe As Integer) As Boolean
    On Error GoTo errLoad
    sp_EtiketResep = True
    Set dbcmd = New ADODB.Command

    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoResep", adVarChar, adParamInput, 15, txtNoResep.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPelayanan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdJenisObat", adChar, adParamInput, 2, IIf(f_KdJenisObat = "", Null, f_KdJenisObat))
        .Parameters.Append .CreateParameter("Signa", adVarChar, adParamInput, 7, f_Signa) 'allow null
        .Parameters.Append .CreateParameter("KdSatuanEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdSatuanEtiket)) = 0, Null, f_KdSatuanEtiket)) 'allow null
        .Parameters.Append .CreateParameter("KdWaktuEtiket", adChar, adParamInput, 2, IIf(Len(Trim(f_KdWaktuEtiket)) = 0, Null, f_KdWaktuEtiket)) 'allow null
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_EtiketResep"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data etiket resep", vbCritical, "Validasi"
            sp_EtiketResep = False
        Else
            Call Add_HistoryLoginActivity("Add_EtiketResep")
        End If
    End With

    Exit Function
errLoad:
    sp_EtiketResep = False
    Call msubPesanError
End Function

Private Sub txtNoResep_LostFocus()
    On Error GoTo errLoad
    If Len(Trim((txtNoResep.Text))) = 0 Then Exit Sub
    strSQL = "SELECT NoResep FROM PemakaianAlkes WHERE (NoResep = '" & txtNoResep.Text & "') AND Year(TglPelayanan) = '" & Year(dtpTglPelayanan.Value) & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "No Resep sudah terpakai, Ganti No Resep", vbExclamation, "Validasi"
        txtNoResep.Text = ""
        txtNoResep.SetFocus
        Call subLoadDataResep(txtNoResep.Text)
        Call subHitungTotal

    End If
    txtNoResep.Text = StrConv(txtNoResep.Text, vbUpperCase)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_TempDetailApotikJual(f_HargaSatuan As Currency, f_KdBarang As String, f_KdAsal As String) As Boolean
    sp_TempDetailApotikJual = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoTemporary", adChar, adParamInput, 3, IIf(Len(Trim(txtNoTemporary.Text)) = 0, Null, Trim(txtNoTemporary.Text)))
        .Parameters.Append .CreateParameter("KdKelompokPasien", adChar, adParamInput, 2, mstrKdJenisPasien)
        .Parameters.Append .CreateParameter("IdPenjamin", adChar, adParamInput, 10, mstrKdPenjaminPasien)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang) 'fgData.TextMatrix(fgData.Row, 2))
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal) 'fgData.TextMatrix(fgData.Row, 12))
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , f_HargaSatuan)
        .Parameters.Append .CreateParameter("NoTemporaryOutput", adChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TemporaryDetailApotikJual"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pengambilan no temporary", vbCritical, "Validasi"
            sp_TempDetailApotikJual = False
        Else
            txtNoTemporary.Text = Trim(.Parameters("NoTemporaryOutput").Value)
            'Call Add_HistoryLoginActivity("Add_TemporaryDetailApotikJual")
        End If
    End With
End Function

Private Sub subHitungTotal()
    On Error GoTo errLoad
    Dim i As Integer

    If fgData.TextMatrix(fgData.Row - 1, 11) = "" Then Exit Sub
    txtTotalBiaya.Text = 0
    txtHutangPenjamin.Text = 0
    txtTanggunganRS.Text = 0
    txtHarusDibayar.Text = 0
    txtTotalDiscount.Text = 0

    With fgData
    
        For i = 1 To IIf(fgData.TextMatrix(fgData.Rows - 1, 2) = "", fgData.Rows - 2, fgData.Rows - 1)
            If .TextMatrix(i, 22) = "" Then .TextMatrix(i, 22) = 0
            txtTotalBiaya.Text = txtTotalBiaya.Text + CDbl(.TextMatrix(i, 11))
            txtHutangPenjamin.Text = txtHutangPenjamin.Text + CDbl(.TextMatrix(i, 19))
            txtTanggunganRS.Text = txtTanggunganRS.Text + CDbl(.TextMatrix(i, 20))
            txtTotalDiscount.Text = txtTotalDiscount.Text + CDbl(.TextMatrix(i, 21))
            txtHarusDibayar.Text = txtHarusDibayar.Text + CDbl(.TextMatrix(i, 22))
        Next i
    End With

    txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, Format(txtTotalBiaya.Text, "#,###"))
    txtHutangPenjamin.Text = IIf(Val(txtHutangPenjamin.Text) = 0, 0, Format(txtHutangPenjamin.Text, "#,###"))
    txtTanggunganRS.Text = IIf(Val(txtTanggunganRS.Text) = 0, 0, Format(txtTanggunganRS.Text, "#,###"))
    txtHarusDibayar.Text = IIf(Val(txtHarusDibayar.Text) = 0, 0, Format(txtHarusDibayar.Text, "#,###"))
    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###"))

    subcurHarusDibayar = txtHarusDibayar.Text

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDataResep(f_NoResep As String)
    On Error GoTo errLoad
    Dim i As Integer
    Dim curHutangPenjamin As Currency
    Dim curHarusDibayar As Currency
    Dim curTanggunganRS As Currency

    strSQL = "SELECT * FROM V_AmbilPemakaianAlkesResep WHERE NoResep = '" & f_NoResep & "'"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        Call subSetGrid
        dtpTglResep.Value = Now
        chkDokterPemeriksa.Value = vbUnchecked
        txtRP.Text = ""
        Exit Sub
    End If

    dtpTglResep.Value = dbRst("TglResep")
    If IsNull(dbRst("IdDokter")) Then
        chkDokterPemeriksa.Value = vbUnchecked
        txtKdDokter.Text = ""
        txtDokter.Text = ""
    Else
        chkDokterPemeriksa.Value = vbChecked
        txtKdDokter.Text = dbRst("IdDokter")
        txtDokter.Text = dbRst("Dokter")
    End If
    dgDokter.Visible = False
    txtRP.Text = dbRst("RuanganResep")

    For i = 0 To dbRst.RecordCount - 1
        'ambil no temporary
        txtKdBarang.Text = dbRst("KdBarang")
        txtKdAsal.Text = dbRst("KdAsal")
        If sp_TempDetailApotikJual(CDbl(dbRst("HargaSatuan")) + CDbl(dbRst("TarifService")), dbRst("KdBarang"), dbRst("KdAsal")) = False Then Exit Sub 'discount
        'ambil hutang penjamin dan tanggungan rs
        strSQL = "SELECT HargaSatuan, JmlHutangPenjamin, JmlTanggunganRS" & _
        " FROM TempDetailApotikJual" & _
        " WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "') AND (KdBarang = '" & dbRst("KdBarang") & "') AND (KdAsal = '" & dbRst("KdAsal") & "')"
        Call msubRecFO(rsB, strSQL)
        If rsB.EOF = True Then
            curHutangPenjamin = 0
            curHarusDibayar = 0
        Else
            curHutangPenjamin = rsB("JmlHutangPenjamin").Value
            curHarusDibayar = rsB("JmlTanggunganRS").Value
        End If

        With fgData
            .TextMatrix(.Rows - 1, 0) = dbRst("ResepKe")
            .TextMatrix(.Rows - 1, 1) = dbRst("JenisObat")
            .TextMatrix(.Rows - 1, 2) = dbRst("KdBarang")
            .TextMatrix(.Rows - 1, 3) = dbRst("NamaBarang")
            .TextMatrix(.Rows - 1, 4) = dbRst("KeKuatan")
            .TextMatrix(.Rows - 1, 5) = dbRst("NamaAsal")
            .TextMatrix(.Rows - 1, 6) = dbRst("SatuanJml")
            .TextMatrix(.Rows - 1, 7) = CDbl(dbRst("HargaSatuan")) + IIf(dbRst("JmlService") = 0, 0, dbRst("TarifService"))
            .TextMatrix(.Rows - 1, 7) = IIf(Val(.TextMatrix(.Rows - 1, 7)) = 0, 0, Format(.TextMatrix(.Rows - 1, 5), "#,###"))
            .TextMatrix(.Rows - 1, 8) = CDbl(0) 'discount
            .TextMatrix(.Rows - 1, 8) = IIf(Val(.TextMatrix(.Rows - 1, 6)) = 0, 0, Format(.TextMatrix(.Rows - 1, 6), "#,###"))
            .TextMatrix(.Rows - 1, 9) = CDbl(dbRst("JmlStok") + dbRst("JmlBarang"))
            .TextMatrix(.Rows - 1, 10) = CDbl(dbRst("JmlBarang"))

            .TextMatrix(.Rows - 1, 11) = ((dbRst("TarifService") * dbRst("JmlService")) + _
            (CDbl(dbRst("HargaSatuan")) * CDbl(.TextMatrix(.Rows - 1, 10))))
            .TextMatrix(.Rows - 1, 11) = IIf(Val(.TextMatrix(.Rows - 1, 11)) = 0, 0, Format(.TextMatrix(.Rows - 1, 11), "#,###"))

            .TextMatrix(.Rows - 1, 12) = dbRst("KdAsal")
            .TextMatrix(.Rows - 1, 13) = dbRst("JenisBarang")
            .TextMatrix(.Rows - 1, 14) = dbRst("TarifService")
            .TextMatrix(.Rows - 1, 15) = dbRst("JmlService")
            .TextMatrix(.Rows - 1, 16) = CDbl(dbRst("HargaSatuan"))
            .TextMatrix(.Rows - 1, 17) = curHutangPenjamin
            .TextMatrix(.Rows - 1, 18) = curTanggunganRS

            .TextMatrix(.Rows - 1, 19) = CDbl(dbRst("JmlBarang")) * curHutangPenjamin
            .TextMatrix(.Rows - 1, 20) = CDbl(dbRst("JmlBarang")) * curTanggunganRS
            .TextMatrix(.Rows - 1, 21) = CDbl(dbRst("JmlBarang")) * CDbl(0) 'discount

            curHarusDibayar = CDbl(.TextMatrix(.Rows - 1, 11)) - (CDbl(.TextMatrix(.Rows - 1, 21)) + _
            CDbl(.TextMatrix(.Rows - 1, 19)) + CDbl(.TextMatrix(.Rows - 1, 120)))
            .TextMatrix(.Rows - 1, 22) = IIf(curHarusDibayar < 0, 0, curHarusDibayar)

            .TextMatrix(.Rows - 1, 23) = txtNoTemporary.Text

            .TextMatrix(.Rows - 1, 24) = CDbl(dbRst("HargaBeli"))
            .TextMatrix(.Rows - 1, 25) = IIf(IsNull(dbRst("KdJenisObat")), "", dbRst("KdJenisObat"))
            .TextMatrix(.Rows - 1, 26) = dbRst("BiayaAdministrasi")

            .Rows = .Rows + 1
            dbRst.MoveNext
            dbConn.Execute "DELETE FROM TempDetailApotikJual WHERE (NoTemporary = '" & Trim(txtNoTemporary.Text) & "')"
        End With
    Next i

    Call subHitungTotal

    dgObatAlkes.Visible = False
    txtJenisBarang.Text = "": txtKdBarang.Text = "": txtKdAsal.Text = "": txtSatuan.Text = "": txtAsalBarang.Text = ""

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

    Select Case fgData.Col
        Case 0
            txtIsi.MaxLength = 2

        Case 3
            txtIsi.MaxLength = 20

        Case 10
            txtIsi.MaxLength = 4
    End Select

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
End Sub

Private Sub subLoadDataCombo(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    s_DcName.Width = fgData.ColWidth(fgData.Col)
    s_DcName.Height = fgData.RowHeight(fgData.Row)

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub subLoadCheck()
    Dim i As Integer
    chkStatusStok.Left = fgData.Left

    For i = 0 To fgData.Col - 1
        chkStatusStok.Left = chkStatusStok.Left + fgData.ColWidth(i)
    Next i
    chkStatusStok.Visible = True
    chkStatusStok.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        chkStatusStok.Top = chkStatusStok.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        chkStatusStok.Top = chkStatusStok.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    chkStatusStok.Width = fgData.ColWidth(fgData.Col)
    chkStatusStok.Height = fgData.RowHeight(fgData.Row)
    chkStatusStok.BackColor = fgData.BackColor

    chkStatusStok.Visible = True

    chkStatusStok.SetFocus
End Sub

Private Sub txtRP_Change()
    On Error GoTo hell
    'Update 15-05-06 JSPRJ
    If Len(Trim(txtRP.Text)) = 0 Then StrKdRP = "": Exit Sub
    strSQL = "Select KdRuangan from Ruangan  where NamaRuangan='" & txtRP.Text & "'"
    Call msubRecFO(rs, strSQL)
    StrKdRP = IIf(IsNull(rs.Fields(0).Value), "", rs.Fields(0).Value)
    Exit Sub
hell:
    Call msubPesanError
End Sub

