VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateBiayaPelayananOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah Biaya Pelayanan Obat & Alkes"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdateBiayaPelayananOA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   12855
   Begin VB.Frame fraKode 
      Caption         =   "Kode2"
      Height          =   1095
      Left            =   4080
      TabIndex        =   44
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtKdAsal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "txtKdAsal"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtKdBarang 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1320
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "txtKdBarang"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtIdPemeriksa 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "txtIdPemeriksa"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtKdRuangan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3120
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "txtKdRuangan"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1095
      Left            =   0
      TabIndex        =   28
      Top             =   960
      Width           =   12855
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   225
         MaxLength       =   10
         TabIndex        =   39
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame4 
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
         Height          =   735
         Left            =   10185
         TabIndex        =   32
         Top             =   240
         Width           =   2535
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   35
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            MaxLength       =   6
            TabIndex        =   34
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            MaxLength       =   6
            TabIndex        =   33
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1440
            TabIndex        =   38
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   37
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            TabIndex        =   36
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8745
         MaxLength       =   9
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3825
         MaxLength       =   50
         TabIndex        =   29
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
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
         Index           =   0
         Left            =   225
         TabIndex        =   43
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
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
         Left            =   8745
         TabIndex        =   42
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
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
         Left            =   1905
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
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
         Left            =   3825
         TabIndex        =   40
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame fraPelayanan 
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
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   12855
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   12
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8280
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   9540
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10680
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddKomponen 
         Caption         =   "+"
         Height          =   375
         Left            =   11880
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdMinKomponen 
         Caption         =   "-"
         Height          =   375
         Left            =   12375
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2640
         Width           =   5415
      End
      Begin VB.TextBox txtTotaltarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtRuanganPelayanan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtPemeriksa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1920
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker dtpTglPendaftaran 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   12648447
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   125173763
         UpDown          =   -1  'True
         CurrentDate     =   38537
      End
      Begin MSDataListLib.DataCombo dcKomponenTarif 
         Height          =   330
         Left            =   5760
         TabIndex        =   14
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   1575
         Left            =   5760
         TabIndex        =   15
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   -2147483633
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpTglPerubahan 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   125173763
         UpDown          =   -1  'True
         CurrentDate     =   38537
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Pelayanan"
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
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
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
         Index           =   2
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Tarif"
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
         Index           =   3
         Left            =   5760
         TabIndex        =   25
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tarif"
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
         Index           =   4
         Left            =   8280
         TabIndex        =   24
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
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
         Left            =   9540
         TabIndex        =   23
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Charge"
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
         Index           =   6
         Left            =   10680
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
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
         Index           =   9
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
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
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Perubahan"
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
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
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
         Index           =   13
         Left            =   10320
         TabIndex        =   18
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Pelayanan"
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
         Index           =   11
         Left            =   2280
         TabIndex        =   17
         Top             =   960
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   450
      Left            =   11160
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   450
      Left            =   9480
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   50
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
      Left            =   11040
      Picture         =   "frmUpdateBiayaPelayananOA.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUpdateBiayaPelayananOA.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Index           =   10
      Left            =   9600
      TabIndex        =   49
      Top             =   2640
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUpdateBiayaPelayananOA.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmUpdateBiayaPelayananOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub cmdAddKomponen_Click()

    On Error GoTo errLoad

    If dcKomponenTarif.MatchedWithList = False Then dcKomponenTarif.SetFocus: Exit Sub
    For i = 1 To fgData.Rows - 1
        If fgData.TextMatrix(i, 5) = dcKomponenTarif.BoundText Then
            fgData.TextMatrix(i, 3) = txtDiscount.Text 'discount
            fgData.TextMatrix(i, 4) = txtCharge.Text  'charge
            Call subHitungTotal
            Exit Sub
        End If
    Next i

    fgData.Rows = fgData.Rows + 1

    fgData.TextMatrix(fgData.Rows - 1, 1) = dcKomponenTarif.Text 'nama komponen
    fgData.TextMatrix(fgData.Rows - 1, 2) = IIf(Val(txtTarif) = 0, 0, Format(txtTarif.Text, "#,###")) 'tarif
    fgData.TextMatrix(fgData.Rows - 1, 3) = txtDiscount.Text 'discount
    fgData.TextMatrix(fgData.Rows - 1, 4) = txtCharge.Text 'charge
    fgData.TextMatrix(fgData.Rows - 1, 5) = dcKomponenTarif.BoundText 'kode komponen tarif

    Call subHitungTotal

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub cmdMinKomponen_Click()

    On Error GoTo errLoad

    If fgData.Rows = 1 Then Exit Sub

    If fgData.Rows = 2 Then
        fgData.TextMatrix(1, 1) = ""
        fgData.TextMatrix(1, 2) = "0"
        fgData.TextMatrix(1, 3) = "0"
        fgData.TextMatrix(1, 4) = "0"
        fgData.Rows = 1
    Else
        fgData.RemoveItem fgData.Row
    End If

    Call subHitungTotal

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub cmdSimpan_Click()

    On Error GoTo errLoad

    If Periksa("text", txtNoPendaftaran, "No Pendaftaran kosong") = False Then Exit Sub
    If subSimpanBackupBiayaPelayananOA = False Then Exit Sub

    For i = 1 To fgData.Rows - 1
        If subSimpanDetailBackupBiayaPelayananOA(fgData.TextMatrix(i, 5), fgData.TextMatrix(i, 3), fgData.TextMatrix(i, 4), fgData.TextMatrix(i, 2)) = False Then Exit Sub
    Next i

    Call Add_HistoryLoginActivity("Add_BackupUpdatingBiayaPelayananOA+Add_DetailBackupUpdatingBiayaPelayananOA")
    cmdSimpan.Enabled = False
    cmdTutup.SetFocus

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub cmdTutup_Click()

    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan perubahan data pelayanan?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Call frmTransaksiPasien.subPemakaianObatAlkes
    Unload Me

End Sub

Private Sub dcDokterPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub dcKomponenTarif_Change()

    On Error GoTo errLoad

    If dcKomponenTarif.MatchedWithList = False Then
        txtTarif.Text = 0
        Exit Sub
    End If

    strSQL = "SELECT TempHargaKomponenObatAlkes.HargaSatuan" & _
    " FROM TempHargaKomponenObatAlkes INNER JOIN KomponenTarif ON TempHargaKomponenObatAlkes.KdKomponen = KomponenTarif.KdKomponen" & _
    " WHERE (TempHargaKomponenObatAlkes.KdBarang = '" & txtKdBarang.Text & "') AND (TempHargaKomponenObatAlkes.KdAsal = '" & txtKdAsal.Text & "') AND (TempHargaKomponenObatAlkes.TglPelayanan = '" & Format(frmTransaksiPasien.dgObatAlkes.Columns("TglPelayanan"), "yyyy/MM/dd HH:mm:ss") & "') AND (TempHargaKomponenObatAlkes.KdRuangan = '" & frmTransaksiPasien.dgObatAlkes.Columns("KdRuangan") & "') AND (TempHargaKomponenObatAlkes.KdKomponen = '" & dcKomponenTarif.BoundText & "') AND (TempHargaKomponenObatAlkes.NoPendaftaran = '" & txtNoPendaftaran.Text & "')"
    Call msubRecFO(rs, strSQL)
    If Not rs.EOF Then
        txtTarif.Text = Format(rs(0).Value, "#,###"): txtTarif.Enabled = False
    Else
        txtTarif.Text = 0: txtTarif.Enabled = True: txtDiscount.Enabled = False: txtCharge.Enabled = False
    End If

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcKomponenTarif_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtTarif.Enabled = True Then
            txtTarif.SetFocus
        Else
            If txtDiscount.Enabled = True Then
                txtDiscount.SetFocus
            Else
                If txtCharge.Enabled = True Then
                    txtCharge.SetFocus
                Else
                    cmdAddKomponen.SetFocus
                End If
            End If
        End If
    End If

End Sub

Private Sub dtpTglPendaftaran_Change()
    dtpTglPendaftaran.MaxDate = Now
End Sub

Private Sub dtpTglPerubahan_Change()
    dtpTglPerubahan.MaxDate = Now
End Sub

Private Sub dtpTglPerubahan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPemeriksa.SetFocus
End Sub

Private Sub fgData_Click()

    If fgData.Row = 0 Then Exit Sub
    dcKomponenTarif.BoundText = fgData.TextMatrix(fgData.Row, 5)
    txtTarif.Text = fgData.TextMatrix(fgData.Row, 2)
    txtDiscount.Text = fgData.TextMatrix(fgData.Row, 3)
    txtCharge.Text = fgData.TextMatrix(fgData.Row, 4)

    txtDiscount.Enabled = True: txtCharge.Enabled = True

End Sub

Private Sub Form_Load()

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subKosong
    Call subSetGrid
    Call subLoadDcSource

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub txtBiayaObatAlkes_KeyPress(KeyAscii As Integer)
    SetKeyPressToNumber KeyAscii
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtCharge_Change()
    If Val(txtCharge.Text) <> 0 Then txtDiscount.Text = 0
End Sub

Private Sub txtCharge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAddKomponen.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtCharge_LostFocus()
    txtCharge = IIf(Val(txtCharge) = 0, 0, Format(txtCharge, "#,###"))
    If Val(txtCharge.Text) > 0 Then
        txtDiscount.Enabled = False
    Else
        txtDiscount.Enabled = True
    End If
End Sub

Private Sub txtDiscount_Change()
    If Val(txtDiscount.Text) <> 0 Then txtCharge.Text = 0
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If txtCharge.Enabled = True Then txtCharge.SetFocus Else cmdAddKomponen.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount = IIf(Val(txtDiscount) = 0, 0, Format(txtDiscount, "#,###"))
    If Val(txtDiscount.Text) > 0 Then
        txtCharge.Enabled = False
    Else
        txtCharge.Enabled = True
    End If
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If frmDaftarPasienRI.optPasAktif.Value = True Then
            strSQL = "Select * from V_DaftarPasienRIAktif WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
        Else
            strSQL = "Select * from V_DaftarPasienRIPindahKamar WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "'"
        End If
        Call msubRecFO(rs, strSQL)
        If rs.EOF Then Call subKosong: Call subSetGrid: Exit Sub

        txtNoCM.Text = rs("NoCM").Value
        txtNamaPasien.Text = rs("Nama Pasien").Value
        txtJK.Text = IIf(rs("JK").Value = "L", "Laki-Laki", "Perempuan")
        txtThn.Text = rs("UmurTahun").Value
        txtBln.Text = rs("UmurBulan").Value
        txtHr.Text = rs("UmurHari").Value

        strSQL = "SELECT * " & _
        " FROM V_UbahPemakaianAlkes" & _
        " WHERE NoPendaftaran = '" & txtNoPendaftaran.Text & "' AND (KdBarang = '" & frmTransaksiPasien.dgObatAlkes.Columns("KdBarang") & "') AND (KdAsal = '" & frmTransaksiPasien.dgObatAlkes.Columns("KdAsal") & "') AND (TglPelayanan = '" & Format(frmTransaksiPasien.dgObatAlkes.Columns("TglPelayanan"), "yyyy/MM/dd HH:mm:ss") & "') AND (KdRuangan = '" & frmTransaksiPasien.dgObatAlkes.Columns("KdRuangan") & "')"
        Call msubRecFO(rs, strSQL)

        dtpTglPendaftaran.Value = rs("TglPelayanan").Value
        txtKdBarang.Text = rs("KdBarang").Value
        txtKdAsal.Text = rs("KdAsal").Value
        txtNamaBarang.Text = Trim(frmTransaksiPasien.dgObatAlkes.Columns("NamaBarang").Value)
        txtRuanganPelayanan.Text = rs("NamaRuangan").Value
        txtKdRuangan.Text = Trim(frmTransaksiPasien.dgObatAlkes.Columns("KdRuangan"))
        txtPemeriksa.Text = rs("Pemeriksa").Value
        txtIdPemeriksa.Text = rs("IdPegawai").Value

        fgData.Rows = rs.RecordCount + 1
        For i = 1 To rs.RecordCount
            fgData.TextMatrix(i, 1) = rs("NamaKomponen").Value
            fgData.TextMatrix(i, 2) = IIf(rs("Harga").Value = 0, 0, Format(rs("Harga").Value, "#,###"))
            fgData.TextMatrix(i, 3) = 0
            fgData.TextMatrix(i, 4) = 0
            fgData.TextMatrix(i, 5) = rs("KdKomponen").Value
            rs.MoveNext
        Next i

        Call subHitungTotal

        cmdSimpan.Enabled = True
        dcKomponenTarif.SetFocus
    End If

End Sub

Private Sub subLoadDcSource()
    Call msubDcSource(dcKomponenTarif, rs, "SELECT KdKomponen, NamaKomponen FROM KomponenTarif where StatusEnabled='1' order by NamaKomponen")
End Sub

Private Sub subKosong()

    txtNoCM.Text = ""
    txtNamaPasien.Text = ""
    txtJK.Text = ""
    txtThn.Text = ""
    txtBln.Text = ""
    txtHr.Text = ""
    dtpTglPendaftaran.Value = Now
    txtNamaBarang.Text = ""
    dcKomponenTarif.BoundText = ""
    txtTarif.Text = "0"
    txtDiscount.Text = "0"
    txtCharge.Text = "0"
    txtTotaltarif.Text = "0"

    txtKdBarang.Text = ""
    txtKdAsal.Text = ""

    dtpTglPerubahan.Value = Now
    txtKeterangan.Text = ""

    'nebeng
    cmdSimpan.Enabled = False
    txtTarif.Enabled = True

End Sub

Private Sub subSetGrid()
    With fgData
        .clear
        .Cols = 6
        .Rows = 1

        .ColWidth(0) = 0
        .ColWidth(1) = 2400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 0

        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter

        .TextMatrix(0, 1) = "Nama Komponen"
        .TextMatrix(0, 2) = "Tarif"
        .TextMatrix(0, 3) = "Discount"
        .TextMatrix(0, 4) = "Charge"
        .TextMatrix(0, 5) = "Kode Komponen Tarif"
    End With
End Sub

Private Sub subHitungTotal()

    On Error GoTo errLoad

    txtTotaltarif = 0

    For i = 1 To fgData.Rows - 1
        'total tarif
        txtTotaltarif.Text = CCur(txtTotaltarif.Text) + _
        IIf(Val(fgData.TextMatrix(i, 2)) = 0, 0, CCur(fgData.TextMatrix(i, 2))) - _
        IIf(Val(fgData.TextMatrix(i, 3)) = 0, 0, CCur(fgData.TextMatrix(i, 3))) + _
        IIf(Val(fgData.TextMatrix(i, 4)) = 0, 0, CCur(fgData.TextMatrix(i, 4)))
    Next i

    txtTotaltarif.Text = IIf(Val(txtTotaltarif) = 0, 0, Format(txtTotaltarif.Text, "#,###"))

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub txtTarif_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtDiscount.Enabled = True Then
            txtDiscount.SetFocus
        Else
            If txtCharge.Enabled = True Then txtCharge.SetFocus Else cmdAddKomponen.SetFocus
        End If
    End If

    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0

End Sub

Private Sub txtTarif_LostFocus()
    txtTarif = IIf(Val(txtTarif) = 0, 0, Format(txtTarif, "#,###"))
End Sub

Private Function subSimpanBackupBiayaPelayananOA() As Boolean

    subSimpanBackupBiayaPelayananOA = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, txtKdRuangan.Text)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("TglUpdate", adDate, adParamInput, , Format(dtpTglPerubahan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 150, IIf(Len(Trim(txtKeterangan.Text)) = 0, "-", Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BackupUpdatingBiayaPelayananOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan saat penyimpanan data backup biaya pelayanan OA", vbCritical, vbOKOnly, "Validasi"
            subSimpanBackupBiayaPelayananOA = False

        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

End Function

Private Function subSimpanDetailBackupBiayaPelayananOA(f_strKdKomponen As String, f_curDiscount As Currency, f_curCharge As Currency, f_curTarif As Currency) As Boolean

    subSimpanDetailBackupBiayaPelayananOA = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, txtKdRuangan.Text)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtpTglPendaftaran.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, txtKdAsal.Text)
        .Parameters.Append .CreateParameter("KdKomponen", adChar, adParamInput, 2, f_strKdKomponen)
        .Parameters.Append .CreateParameter("JmlDiscount", adCurrency, adParamInput, , CCur(f_curDiscount))
        .Parameters.Append .CreateParameter("Tarif", adCurrency, adParamInput, , CCur(f_curTarif))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DetailBackupUpdatingBiayaPelayananOA"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan saat penyimpanan data detail backup biaya pelayanan OA", vbCritical, vbOKOnly, "Validasi"
            subSimpanDetailBackupBiayaPelayananOA = False

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

End Function

