VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmUbahKasusPenyakitPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Ubah Kasus Penyakit Pasien"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUbahKasusPenyakitPasien.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9885
   Begin VB.TextBox txtNamaFormPengirim 
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
      Left            =   0
      MaxLength       =   50
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   3240
      Width           =   9855
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
         Left            =   8160
         TabIndex        =   10
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
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data SMF (Kasus Penyakit)"
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
      Left            =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   9855
      Begin MSDataListLib.DataCombo dcSMFLama 
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   645
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo dcSMFBaru 
         Height          =   360
         Left            =   5160
         TabIndex        =   8
         Top             =   645
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
         Caption         =   "SMF (Kasus Penyakit) BARU"
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
         TabIndex        =   23
         Top             =   360
         Width           =   2370
      End
      Begin VB.Label lblRegistrasiInap 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit) LAMA"
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2385
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
      TabIndex        =   11
      Top             =   960
      Width           =   9855
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
         Left            =   7080
         TabIndex        =   12
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         Left            =   5760
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
         Width           =   1455
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
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   19
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
         Left            =   5760
         TabIndex        =   18
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
         TabIndex        =   17
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
         Left            =   3240
         TabIndex        =   16
         Top             =   360
         Width           =   1110
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   25
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
      Left            =   8160
      Picture         =   "frmUbahKasusPenyakitPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUbahKasusPenyakitPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUbahKasusPenyakitPasien.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmUbahKasusPenyakitPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intJmlDokter As Integer

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If Periksa("datacombo", dcSMFBaru, "SMF / Kasus penyakit kosong") = False Then Exit Sub
    If sp_KasusPenyakitPasien() = False Then Exit Sub
    MsgBox "penyimpanan data berhasil", vbInformation, "Informasi"
    cmdSimpan.Enabled = False

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcSMFBaru_KeyPress(KeyAscii As Integer)

    On Error GoTo errLoad
    If KeyAscii = 13 Then
        If Len(Trim(dcSMFBaru.Text)) = 0 Then Exit Sub
        If dcSMFBaru.MatchedWithList = True Then
            If cmdSimpan.Enabled = False Then
                cmdTutup.SetFocus
            Else
                cmdSimpan.SetFocus
            End If
        End If
        strSQL = "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubinstalasiRuangan WHERE (NamaSubInstalasi LIKE '%" & dcSMFBaru.Text & "%') AND (KdRuangan= '" & mstrKdRuangan & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcSMFBaru.BoundText = rs(0).Value
        dcSMFBaru.Text = rs(1).Value
    End If

    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDcSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtNamaFormPengirim.Text = "frmDaftarPasienRI" Then frmDaftarPasienRI.Enabled = True: frmDaftarPasienRI.cmdCari_Click
End Sub

Private Sub subDcSource()

    On Error GoTo errLoad
    Call msubDcSource(dcSMFLama, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubinstalasiRuangan WHERE (KdRuangan= '" & mstrKdRuangan & "') and StatusEnabled='1'")
    Call msubDcSource(dcSMFBaru, rs, "SELECT KdSubInstalasi, NamaSubInstalasi FROM V_SubinstalasiRuangan WHERE (KdRuangan= '" & mstrKdRuangan & "') and StatusEnabled='1'")
    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Public Sub txtNoPendaftaran_KeyPress(KeyAscii As Integer)

    On Error GoTo hell
    If KeyAscii = 13 Then
        strSQL = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,KdKamar" & _
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

        dcSMFLama.BoundText = rs("KdSubInstalasi")
        dcSMFBaru.SetFocus
    End If
    Exit Sub

hell:
    Call msubPesanError

End Sub

Private Function sp_KasusPenyakitPasien() As Boolean

    On Error GoTo errLoad

    sp_KasusPenyakitPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdSubInstalasiBaru", adChar, adParamInput, 3, dcSMFBaru.BoundText)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(mdTglMasuk, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_KasusPenyakitPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan saat update kasus penyakit pasien", vbCritical, "Validasi"
            sp_KasusPenyakitPasien = False
        Else
            Call Add_HistoryLoginActivity("Update_KasusPenyakitPasien")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function
errLoad:
    sp_KasusPenyakitPasien = False
    Call msubPesanError

End Function

