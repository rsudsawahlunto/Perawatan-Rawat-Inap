VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmUbahRegistrasiPasienRI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah Registrasi Pasien"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahRegistrasiPasienRI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   11265
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   2055
      Left            =   -3720
      TabIndex        =   34
      Top             =   -1440
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtIdDokterBaru 
         Height          =   495
         Left            =   2880
         TabIndex        =   41
         Text            =   "IdDokterBaru"
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtKdKelasPelLama 
         Height          =   495
         Left            =   1440
         TabIndex        =   40
         Text            =   "KdKelasPelLama"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTglMasuk 
         Height          =   495
         Left            =   1440
         TabIndex        =   39
         Text            =   "TglMasuk"
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtNoBedLama 
         Height          =   495
         Left            =   0
         TabIndex        =   38
         Text            =   "NoBedLama"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtNoKamarLama 
         Height          =   495
         Left            =   380
         TabIndex        =   37
         Text            =   "NoKamarLama"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtKdSubInstalasiLama 
         Height          =   495
         Left            =   0
         TabIndex        =   36
         Text            =   "KdSubInstalasiLama"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNoPakai 
         Height          =   495
         Left            =   0
         TabIndex        =   35
         Text            =   "nopakai"
         Top             =   0
         Width           =   1215
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
      Height          =   2415
      Left            =   2040
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   1935
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
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
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   30
      Top             =   3960
      Width           =   11175
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
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
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
         Left            =   7560
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Registrasi Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   24
      Top             =   2040
      Width           =   11175
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5160
         TabIndex        =   12
         Top             =   1365
         Width           =   5775
      End
      Begin MSDataListLib.DataCombo dcKelasPelayanan 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   645
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcKelasKamar 
         Height          =   360
         Left            =   3240
         TabIndex        =   8
         Top             =   645
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoKamar 
         Height          =   360
         Left            =   6720
         TabIndex        =   9
         Top             =   645
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcNoBed 
         Height          =   360
         Left            =   9720
         TabIndex        =   10
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcInstalasi 
         Height          =   360
         Left            =   210
         TabIndex        =   11
         Top             =   1365
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Penanggung Jawab"
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
         Left            =   5160
         TabIndex        =   33
         Top             =   1080
         Width           =   2505
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
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
         Left            =   210
         TabIndex        =   29
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Kamar"
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
         Left            =   6720
         TabIndex        =   28
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "No. Bed"
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
         Left            =   9720
         TabIndex        =   27
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Kamar"
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
         Left            =   3240
         TabIndex        =   26
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
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
         Left            =   210
         TabIndex        =   25
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
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
      TabIndex        =   15
      Top             =   960
      Width           =   11175
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
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
         Left            =   8400
         TabIndex        =   16
         Top             =   240
         Width           =   2655
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            TabIndex        =   6
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtBln 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            TabIndex        =   5
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            TabIndex        =   4
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            TabIndex        =   19
            Top             =   360
            Width           =   195
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
            TabIndex        =   18
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label7 
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
            TabIndex        =   17
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   7080
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
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
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
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
         Left            =   7080
         TabIndex        =   22
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
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
         Left            =   1680
         TabIndex        =   21
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
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   1110
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   42
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
      Picture         =   "frmUbahRegistrasiPasienRI.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahRegistrasiPasienRI.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahRegistrasiPasienRI.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmUbahRegistrasiPasienRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intJmlDokter As Integer

Private Sub cmdSimpan_Click()

    On Error GoTo errLoad

    If Periksa("datacombo", dcKelasPelayanan, "Kelas pelayanan tidak terdaftar") = False Then Exit Sub
    If Periksa("datacombo", dcKelasKamar, "Kelas kamar tidak terdaftar") = False Then Exit Sub
    If Periksa("datacombo", dcNoKamar, "No kamar tidak terdaftar") = False Then Exit Sub
    If Periksa("datacombo", dcNoBed, "No bed tidak terdaftar") = False Then Exit Sub
    If Periksa("datacombo", dcInstalasi, "SMF / Kasus penyakit tidak terdaftar") = False Then Exit Sub

    strSQL = "SELECT StatusBed FROM StatusBed WHERE (KdKamar = '" & dcNoKamar.BoundText & "' ) AND (NoBed = '" & dcNoBed.BoundText & "')"
    Call msubRecFO(rs, strSQL)
    If UCase(rs(0).Value) = "I" Then
        MsgBox "Status bed sudah terisi", vbExclamation, "Validasi"
        Call msubDcSource(dcNoBed, rsB, "SELECT DISTINCT NoBed, NoBed AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') AND KdKamar = '" & dcNoKamar.BoundText & "'")
        dcNoBed.BoundText = ""
        Exit Sub
    End If

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, txtNoPakai.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasiLama", adChar, adParamInput, 3, txtKdSubInstalasiLama.Text)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdKamarLama", adChar, adParamInput, 4, txtNoKamarLama.Text)
        .Parameters.Append .CreateParameter("NoBedLama", adChar, adParamInput, 2, txtNoBedLama.Text)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(txtTglMasuk.Text, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKelasPelLama", adChar, adParamInput, 2, txtKdKelasPelLama.Text)
        .Parameters.Append .CreateParameter("KdSubInstalasiBaru", adChar, adParamInput, 3, dcInstalasi.BoundText)
        .Parameters.Append .CreateParameter("IdDokterBaru", adChar, adParamInput, 10, IIf(Len(Trim(txtIdDokterBaru.Text)) = 0, Null, txtIdDokterBaru.Text)) 'allow null
        .Parameters.Append .CreateParameter("KdKelasBaru", adChar, adParamInput, 2, dcKelasKamar.BoundText)
        .Parameters.Append .CreateParameter("KdKamarBaru", adChar, adParamInput, 4, dcNoKamar.BoundText)
        .Parameters.Append .CreateParameter("NoBedBaru", adChar, adParamInput, 2, dcNoBed.BoundText)
        .Parameters.Append .CreateParameter("KdKelasPelBaru", adChar, adParamInput, 2, dcKelasPelayanan.BoundText) 'kelas pelayanan
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_PasienMasukKamar"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam update pasien masuk kamar", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Update_PasienMasukKamar")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    cmdSimpan.Enabled = False

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcInstalasi.Text)) = 0 Then txtDokter.SetFocus: Exit Sub
        If dcInstalasi.MatchedWithList = True Then txtDokter.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubinstalasiRuangan WHERE NamaSubInstalasi LIKE '%" & dcInstalasi.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcInstalasi.BoundText = dbRst(0).Value
        dcInstalasi.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelasKamar_Change()

    On Error GoTo errLoad
    Call msubDcSource(dcNoKamar, rs, "SELECT DISTINCT KdKamar, NoKamar AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') ANd KdKelas = '" & dcKelasKamar.BoundText & "' and  Expr1='1'")
    dcNoKamar.Text = ""
    Call msubDcSource(dcNoBed, rsB, "SELECT DISTINCT NoBed, NoBed AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') And KdKelas = '" & dcKelasKamar.BoundText & "' and  Expr2='1'")
    dcNoBed.Text = ""
    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcKelasKamar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNoKamar.SetFocus
End Sub

Private Sub dcKelasKamar_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKelasKamar.Text)) = 0 Then dcNoKamar.SetFocus: Exit Sub
        If dcKelasKamar.MatchedWithList = True Then dcNoKamar.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE Kelas LIKE '%" & dcKelasKamar.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKelasKamar.BoundText = dbRst(0).Value
        dcKelasKamar.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKelasPelayanan_Change()

    On Error GoTo errLoad
    Call msubDcSource(dcKelasKamar, dbRst, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') AND (KdKelas = '" & dcKelasPelayanan.BoundText & "') and StatusEnabled='1'")
    If dbRst.EOF = False Then dcKelasKamar.BoundText = dbRst(0).Value
    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub dcKelasPelayanan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelasKamar.SetFocus
End Sub

Private Sub dcKelasPelayanan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKelasPelayanan.Text)) = 0 Then dcKelasKamar.SetFocus: Exit Sub
        If dcKelasPelayanan.MatchedWithList = True Then dcKelasKamar.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE Kelas LIKE '%" & dcKelasPelayanan.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcKelasPelayanan.BoundText = dbRst(0).Value
        dcKelasPelayanan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoBed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcInstalasi.SetFocus
End Sub

Private Sub dcNoBed_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If Len(Trim(dcNoBed.Text)) = 0 Then dcInstalasi.SetFocus: Exit Sub
        If dcNoBed.MatchedWithList = True Then dcInstalasi.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT DISTINCT NoBed, NoBed As Alias FROM V_KamarRawatInap WHERE NoBed LIKE '%" & dcNoBed.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcNoBed.BoundText = dbRst(0).Value
        dcNoBed.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcNoKamar_Change()
    On Error GoTo hell
    Call msubDcSource(dcNoBed, rsB, "SELECT DISTINCT NoBed, NoBed AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') AND KdKamar = '" & dcNoKamar.BoundText & "' and  Expr2='1'")
    dcNoBed.Text = ""
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcNoKamar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNoBed.SetFocus
End Sub

Private Sub dcNoKamar_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then
        If Len(Trim(dcNoKamar.Text)) = 0 Then dcNoBed.SetFocus: Exit Sub
        If dcNoBed.MatchedWithList = True Then dcNoBed.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT DISTINCT KdKamar, NoKamar AS Alias FROM V_KamarRawatInap WHERE NoKamar LIKE '%" & dcNoKamar.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcNoKamar.BoundText = dbRst(0).Value
        dcNoKamar.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)

    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If intJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(0).Value
        txtIdDokterBaru.Text = dgDokter.Columns(1).Value
        If txtIdDokterBaru = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        Me.Height = 5235
        fraDokter.Visible = False
        cmdSimpan.SetFocus
    End If
errLoad:

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDcSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnFormDaftarPasienRI = True Then
        frmDaftarPasienRI.Enabled = True
        Call frmDaftarPasienRI.cmdCari_Click
    End If
End Sub

Private Sub txtDokter_Change()
    txtIdDokterBaru.Text = ""
    Call subLoadDokter
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()

    On Error GoTo errLoad
    strSQL = "SELECT NamaDokter AS [Nama Dokter], KodeDokter AS [Kode Dokter], JK, Jabatan FROM V_DaftarDokter WHERE NamaDokter LIKE '%" & txtDokter.Text & "%'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDokter = rs.RecordCount
    Set dgDokter.DataSource = rs
    With dgDokter
        .Columns(0).Width = 3000
        .Columns(1).Width = 0
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    Me.Height = 6900
    fraDokter.Visible = True
    Exit Sub
errLoad:

End Sub

Private Sub subDcSource()

    On Error GoTo errLoad
    Call msubDcSource(dcKelasPelayanan, rs, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE (Ruangan = '" & mstrNamaRuangan & "') and StatusEnabled='1'")
    Call msubDcSource(dcKelasKamar, rs, "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRegRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') and StatusEnabled='1'")
    Call msubDcSource(dcNoKamar, rs, "SELECT DISTINCT KdKamar, NoKamar AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') and  Expr1='1'")
    Call msubDcSource(dcNoBed, rs, "SELECT DISTINCT NoBed, NoBed AS Alias FROM V_KamarRawatInap WHERE (KdRuangan = '" & mstrKdRuangan & "') and Expr2='1'")
    Call msubDcSource(dcInstalasi, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubinstalasiRuangan WHERE (KdRuangan= '" & mstrKdRuangan & "') and StatusEnabled='1'")
    Exit Sub

errLoad:
    Call msubPesanError

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
    If KeyAscii = 27 Then
        fraDokter.Visible = False
        Me.Height = 5235
    End If
End Sub

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)

    On Error GoTo hell
    If KeyAscii = 13 Then
        strSQL = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,DokterPenanggungJawab,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,KdKamar" & _
        " from V_DaftarPasienRIAktif " & _
        " where NoPendaftaran like '" & txtnopendaftaran.Text & "'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF Then Exit Sub
        txtnocm.Text = rs("NoCM")
        txtNamaPasien.Text = rs("Nama Pasien")
        If rs("JK") = "L" Then
            txtJK.Text = "Laki-Laki"
        Else
            txtJK.Text = "Perempuan"
        End If
        txtThn.Text = rs("UmurTahun")
        txtBln.Text = rs("UmurBulan")
        txtHr.Text = rs("UmurHari")

        dcKelasPelayanan.BoundText = rs("KdKelas")
        dcNoKamar.BoundText = rs("KdKamar")
        dcNoBed.BoundText = rs("NoBed")
        dcInstalasi.BoundText = rs("KdSubInstalasi")

        txtNoPakai.Text = rs("NoPakai")
        txtKdSubInstalasiLama.Text = rs("KdSubInstalasi")
        txtNoKamarLama.Text = rs("KdKamar")
        txtNoBedLama.Text = rs("NoBed")
        txtTglMasuk.Text = rs("TglMasuk")
        txtKdKelasPelLama.Text = rs("KdKelas")
        If (IsNull(rs.Fields("DokterPenanggungJawab").Value)) Then
            txtDokter.Text = ""
        Else
            txtDokter.Text = rs("DokterPenanggungJawab")
        End If
        fraDokter.Visible = False
        Me.Height = 5235
        dcKelasPelayanan.SetFocus
    End If
    Exit Sub

hell:
    Call msubPesanError

End Sub

