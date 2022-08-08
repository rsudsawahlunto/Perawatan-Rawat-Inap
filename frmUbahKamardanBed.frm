VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUbahKamardanBed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah No. Kamar dan No. Bed"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahKamardanBed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   11295
   Begin VB.TextBox txtKdRuanganAsal 
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Text            =   "txtKdRuanganAsal"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   3960
      Width           =   11295
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   9360
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pasien Masuk Kamar"
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
      TabIndex        =   23
      Top             =   2040
      Width           =   11295
      Begin VB.TextBox txtNoBedLama 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   40
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNoKamLama 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   39
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtRuangPerawatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpTglMasuk 
         Height          =   330
         Left            =   3960
         TabIndex        =   0
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   127008771
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin MSDataListLib.DataCombo dcNoKam 
         Height          =   330
         Left            =   8640
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo dcNoBed 
         Height          =   330
         Left            =   8640
         TabIndex        =   4
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
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
      Begin MSDataListLib.DataCombo dcKelasPK 
         Height          =   330
         Left            =   6240
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo dcKelasKamar 
         Height          =   330
         Left            =   8520
         TabIndex        =   2
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker dtpTglPerubahan 
         Height          =   330
         Left            =   3960
         TabIndex        =   37
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   127008771
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Perubahan"
         Height          =   210
         Left            =   3960
         TabIndex        =   38
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "No. Bed Baru"
         Height          =   210
         Left            =   6240
         TabIndex        =   36
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Kamar Baru"
         Height          =   210
         Left            =   6240
         TabIndex        =   35
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ruang Perawatan"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Kamar"
         Height          =   210
         Left            =   8520
         TabIndex        =   29
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Masuk"
         Height          =   210
         Left            =   3960
         TabIndex        =   28
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Kelas Pelayanan"
         Height          =   210
         Left            =   6240
         TabIndex        =   26
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Kamar Lama"
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   1140
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Bed Lama"
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   1500
         Width           =   1140
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
      Width           =   11295
      Begin VB.TextBox txtNoPakai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   32
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
         Left            =   8760
         TabIndex        =   12
         Top             =   360
         Width           =   2415
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   240
            Width           =   375
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
            TabIndex        =   18
            Top             =   270
            Width           =   150
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
            TabIndex        =   17
            Top             =   270
            Width           =   210
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
            TabIndex        =   16
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   9
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pakai"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   7440
         TabIndex        =   19
         Top             =   360
         Width           =   1065
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
      Left            =   9480
      Picture         =   "frmUbahKamardanBed.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahKamardanBed.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahKamardanBed.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmUbahKamardanBed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrKdDokterPenanggungjawab As String
Dim intJmlDokter As Integer
Dim strFilter As String

Private Sub cmdSimpan_Click()

    On Error GoTo hell
    If funcCekValidasi = False Then Exit Sub
    cmdSimpan.Enabled = False
    strSQL = "SELECT StatusBed FROM StatusBed WHERE (KdKamar = '" & dcNoKam.BoundText & "') AND (NoBed = '" & dcNoBed.BoundText & "')"
    Call msubRecFO(rs, strSQL)
    If UCase(rs(0).Value) = "I" Then
        MsgBox "No bed sudah terpakai", vbExclamation, "Validasi"
        strSQL = "SELECT DISTINCT NoBed FROM V_KamarRawatInap WHERE KdRuangan='" _
        & mstrKdRuangan & "' AND KdKelas='" & dcKelasKamar.BoundText & "' AND " _
        & "NoKamar='" & dcNoKam.Text & "' and StatusBed='K'"
        Call msubDcSource(dcNoBed, rs, strSQL)
        Exit Sub
    End If
    Call subSaveUbahKamardanBed
    Exit Sub

hell:
    Call msubPesanError

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcKelasKamar_Change()

    On Error GoTo hell
    strSQL = "SELECT DISTINCT NoKamar,KdKamar FROM V_KamarRawatInap WHERE KdRuangan='" _
    & mstrKdRuangan & "' AND KdKelas='" & dcKelasKamar.BoundText & "' and Expr1='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcNoKam.RowSource = rs
    dcNoKam.ListField = rs.Fields(0).Name
    dcNoKam.BoundColumn = rs.Fields(1).Name
    Set rs = Nothing
    dcNoKam.Text = ""
    dcNoBed.Text = ""
    Exit Sub

hell:
    Call msubPesanError

End Sub

Private Sub dcKelasKamar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglPerubahan.SetFocus
End Sub

Private Sub dcKelasPK_Change()

    On Error GoTo errLoad
    strSQL = "SELECT DISTINCT KdKelas,Kelas FROM V_KamarRawatInap WHERE KdRuangan='" & mstrKdRuangan & "' AND KdKelas IN ('" & dcKelasPK.BoundText & "','04') and StatusEnabled='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcKelasKamar.RowSource = rs
    dcKelasKamar.BoundColumn = rs.Fields(0).Name
    dcKelasKamar.ListField = rs.Fields(1).Name
    dcKelasKamar.BoundText = rs.Fields(0).Value
    Set rs = Nothing
    dcNoKam.Text = ""
    dcNoBed.Text = ""
    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub dcKelasPK_GotFocus()

    On Error GoTo errLoad
    Dim tempKode As String

    tempKode = dcKelasPK.BoundText
    strSQL = "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRawatInap  WHERE KdRuangan='" & mstrKdRuangan & "' AND KdKelas = '04' and StatusEnabled='1'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        strSQL = "SELECT DISTINCT KdKelas,Kelas FROM V_KamarRawatInap WHERE KdKelas <> '04' and StatusEnabled='1'"
    Else
        strSQL = "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRawatInap  WHERE KdRuangan='" & mstrKdRuangan & "' and StatusEnabled='1'"
    End If
    Call msubDcSource(dcKelasPK, rs, strSQL)
    dcKelasPK.BoundText = tempKode

    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub dcKelasPK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcKelasKamar.SetFocus
End Sub

Private Sub dcNoBed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dcNoKam_Change()

    On Error GoTo hell
    strSQL = "SELECT DISTINCT NoBed FROM V_KamarRawatInap WHERE KdRuangan='" _
    & mstrKdRuangan & "' AND KdKelas='" & dcKelasKamar.BoundText & "' AND " _
    & "NoKamar='" & dcNoKam.Text & "' and StatusBed='K' and Expr2='1'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dcNoBed.RowSource = rs
    dcNoBed.ListField = rs.Fields(0).Name
    Set rs = Nothing
    dcNoBed.Text = ""
    Exit Sub

hell:
    Call msubPesanError

End Sub

Private Sub dcNoKam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcNoBed.SetFocus
End Sub

Private Sub dtpTglMasuk_Change()
    dtpTglMasuk.MaxDate = Now
End Sub

Private Sub dtpTglMasuk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcKelasPK.SetFocus
End Sub

Private Sub dtpTglPerubahan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcNoKam.SetFocus
End Sub

Private Sub Form_Load()

    On Error GoTo hell
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglMasuk.Value = Now
    dtpTglPerubahan.Value = Now
    strSQL = "Select IdDokter FROM RegistrasiRI where NoPendaftaran='" _
    & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        mstrKdDokterPenanggungjawab = ""
    Else
        mstrKdDokterPenanggungjawab = rs.Fields(0).Value
    End If
    Set rs = Nothing

    strSQL = "SELECT DISTINCT KdKelas, Kelas FROM V_KamarRawatInap " & _
    " WHERE KdRuangan='" & mstrKdRuangan & "' And KdKelas <> '04'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then Exit Sub
    If rs("KdKelas").Value = "04" Then
        strSQL = "SELECT DISTINCT KdKelas,Kelas FROM V_KamarRawatInap WHERE KdKelas <> '04' and StatusEnabled='1'"
    End If
    Call msubDcSource(dcKelasPK, rs, strSQL)
    Exit Sub

hell:
    Call msubPesanError

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRI.Enabled = True
    Call frmDaftarPasienRI.cmdCari_Click
End Sub

'untuk mencek validasi
Private Function funcCekValidasi() As Boolean
    If dcKelasPK.Text = "" Then
        MsgBox "Kelas pelayanan pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcKelasPK.SetFocus
        Exit Function
    End If
    If dcNoKam.Text = "" Then
        MsgBox "No Kamar pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcNoKam.SetFocus
        Exit Function
    End If
    If dcNoBed.Text = "" Then
        MsgBox "No bed pasien harus diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        dcNoBed.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'Store procedure untuk mengisi data pasien masuk kamar
Private Sub sp_UpdateKamarRI(f_StatusKeluar As String)
    On Error GoTo errLoad
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, txtNoPakai.Text)
        .Parameters.Append .CreateParameter("TglUpdate", adDate, adParamInput, , Format(dtpTglPerubahan.Value, "yyyy-MM-dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("KdKamarBaru", adChar, adParamInput, 4, dcNoKam.BoundText)
        .Parameters.Append .CreateParameter("NoBedBaru", adChar, adParamInput, 2, dcNoBed.BoundText)
        .Parameters.Append .CreateParameter("StatusKeluar", adChar, adParamInput, 2, f_StatusKeluar)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_KamarPasienRI"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan data pasien masuk kamar", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("Update_KamarPasienRI")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Sub
errLoad:
    Call msubPesanError
    cmdSimpan.Enabled = True
End Sub

'untuk enable/disable control2
Private Sub subDisableControl(blnStatus As Boolean)
    dtpTglMasuk.Enabled = blnStatus
    dcKelasPK.Enabled = blnStatus
    dcKelasKamar.Enabled = blnStatus
    dcNoKam.Enabled = blnStatus
    dcNoBed.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

'untuk save pasien ubah kamar dan bed
Public Sub subSaveUbahKamardanBed()
    MousePointer = vbHourglass
    Call sp_UpdateKamarRI("T")
    Call subDisableControl(False)
    MousePointer = vbDefault
End Sub

Private Sub optTidak_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub optYa_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

