VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPasienPulang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pasien Pulang"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasienPulang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   10710
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   32
      Top             =   3120
      Width           =   10695
      Begin VB.CommandButton cmdPasienMati 
         Caption         =   "Pasien Meninggal"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8640
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pasien Pulang"
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
      TabIndex        =   23
      Top             =   2040
      Width           =   10695
      Begin VB.TextBox txtPenerimaPasien 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtLamaDirawat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dcStatusPulang 
         Height          =   330
         Left            =   5760
         TabIndex        =   2
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSComCtl2.DTPicker dtpTglKeluar 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   126812163
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin MSDataListLib.DataCombo dcKondisiPulang 
         Height          =   330
         Left            =   7920
         TabIndex        =   3
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi Pulang"
         Height          =   210
         Left            =   7920
         TabIndex        =   29
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Penerima Pasien"
         Height          =   210
         Left            =   3480
         TabIndex        =   28
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lama Dirawat"
         Height          =   210
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status Pulang"
         Height          =   210
         Left            =   5760
         TabIndex        =   26
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Keluar"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   7
      Top             =   960
      Width           =   10695
      Begin VB.TextBox txtTglMasuk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   30
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   17
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame5 
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
         Height          =   580
         Left            =   6120
         TabIndex        =   8
         Top             =   360
         Width           =   2415
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtHari 
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
            TabIndex        =   9
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "thn"
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
            Left            =   555
            TabIndex        =   14
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "bln"
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
            Left            =   1350
            TabIndex        =   13
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
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
            Left            =   2130
            TabIndex        =   12
            Top             =   270
            Width           =   150
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Masuk"
         Height          =   210
         Left            =   8640
         TabIndex        =   31
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   4920
         TabIndex        =   22
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPasienPulang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmPasienPulang.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPasienPulang.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPasienPulang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPasienMati_Click()
    frmPasienMeninggal.Show
End Sub

Private Sub cmdSimpan_Click()
    If funcCekValidasi = False Then Exit Sub
    If dcKondisiPulang.BoundText = "04" Or dcKondisiPulang.BoundText = "05" Then
        mblnPsnMati = True
    Else
        mblnPsnMati = False
    End If
    If mblnPsnMati = True Then
        With frmPasienMeninggal
            .Show
            .txtNoPendaftaran.Text = mstrNoPen
            .txtNoCM.Text = mstrNoCM
            .txtNamaPasien.Text = txtNamaPasien.Text
            .txtSex.Text = txtSex.Text
            .txtThn.Text = txtThn.Text
            .txtBln.Text = txtBln.Text
            .txtHari.Text = txtHari.Text
        End With
        Me.Enabled = False
        Exit Sub
    End If
    Call subSavePsnPulang
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKondisiPulang_GotFocus()
    strSQL = "SELECT KdKondisiPulang,KondisiPulang FROM KondisiPulang where KdKondisiPulang<>'09' and KdKondisiPulang<>'10'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcKondisiPulang.RowSource = rs
    dcKondisiPulang.BoundColumn = rs.Fields(0).Name
    dcKondisiPulang.ListField = rs.Fields(1).Name
    Set rs = Nothing
End Sub

Private Sub dcStatusPulang_GotFocus()
    strSQL = "SELECT KdStatusPulang,StatusPulang FROM StatusPulang"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcStatusPulang.RowSource = rs
    dcStatusPulang.BoundColumn = rs.Fields(0).Name
    dcStatusPulang.ListField = rs.Fields(1).Name
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglKeluar.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRI.Enabled = True
End Sub

'untuk mencek validasi
Private Function funcCekValidasi() As Boolean
    If txtPenerimaPasien.Text = "" Then
        MsgBox "Penerima pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtPenerimaPasien.SetFocus
        Exit Function
    End If
    If dcStatusPulang.Text = "" Then
        MsgBox "Status pulang pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcStatusPulang.SetFocus
        Exit Function
    End If
    If dcKondisiPulang.Text = "" Then
        MsgBox "Kondisi pulang pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcKondisiPulang.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi data pasien pulang
Private Sub sp_PasienPulang(ByVal adoCommand As ADODB.Command)
    With adoCommand

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("TglPulang", adDate, adParamInput, , Format(dtpTglKeluar.Value, "yyyy-MM-dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("NamaPenerima", adVarChar, adParamInput, 30, txtPenerimaPasien.Text)
        .Parameters.Append .CreateParameter("KdKondisiPulang", adChar, adParamInput, 2, dcKondisiPulang.BoundText)
        .Parameters.Append .CreateParameter("KdStatusPulang", adChar, adParamInput, 2, dcStatusPulang.BoundText)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, noidpegawai)
        .Parameters.Append .CreateParameter("OutputLamaRawat", adInteger, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("KdRuanganLogin", adChar, adParamInput, 3, mstrKdRuangan)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienRIPulang"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan data pasien pulang", vbCritical, "Validasi"
        Else
            If Not IsNull(.Parameters("OutputLamaRawat").Value) Then _
                txtLamaDirawat.Text = .Parameters("OutputLamaRawat").Value
                MsgBox "Penyimpanan data pasien pulang sukses", vbExclamation, "Validasi"
                Call Add_HistoryLoginActivity("Add_PasienRIPulang")
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
        Exit Sub
End Sub

'untuk enable/disable control2
Private Sub subDisableControl(blnStatus As Boolean)
    dtpTglKeluar.Enabled = blnStatus
    txtPenerimaPasien.Enabled = blnStatus
    dcStatusPulang.Enabled = blnStatus
    dcKondisiPulang.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

'untuk save pasien keluar kamar
Public Sub subSavePsnPulang()
    Call sp_PasienPulang(dbcmd)
    frmDaftarPasienRI.Enabled = True
    Call frmDaftarPasienRI.optPasNonAktif_GotFocus
    Call subDisableControl(False)
End Sub

