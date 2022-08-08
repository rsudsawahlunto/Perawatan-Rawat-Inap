VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRencanaPindahPulangPasien 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Rencana Pindah Pulang Pasien"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13650
   Icon            =   "frmRencanaPindahPulangPasien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   13650
   Begin VB.TextBox TxtNoPakai 
      Height          =   375
      Left            =   9000
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtKdRuanganAsal 
      Height          =   375
      Left            =   6480
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtNoOrder 
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Pasien"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   13455
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   14
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         MaxLength       =   9
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
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
         Left            =   9360
         TabIndex        =   4
         Top             =   240
         Width           =   2775
         Begin VB.TextBox txtThn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            MaxLength       =   6
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtHr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "thn"
            Height          =   210
            Left            =   720
            TabIndex        =   10
            Top             =   285
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "bln"
            Height          =   210
            Left            =   1560
            TabIndex        =   9
            Top             =   285
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "hr"
            Height          =   210
            Left            =   2400
            TabIndex        =   8
            Top             =   285
            Width           =   165
         End
      End
      Begin VB.Label lblNamaPasien 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblJnsKlm 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7800
         TabIndex        =   15
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   13455
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSimpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10320
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   13455
      Begin MSDataListLib.DataCombo dcStatusPulang 
         Height          =   330
         Left            =   9120
         TabIndex        =   24
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataListLib.DataCombo dcKondisiPulang 
         Height          =   330
         Left            =   11160
         TabIndex        =   23
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo dcStatusKeluar 
         Height          =   330
         Left            =   2640
         TabIndex        =   21
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DcRuanganTujuan 
         Height          =   330
         Left            =   4560
         TabIndex        =   20
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo dcKelas 
         Height          =   330
         Left            =   7200
         TabIndex        =   22
         Top             =   600
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
      Begin MSComCtl2.DTPicker dtpTglRencanaKeluar 
         Height          =   330
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   106037251
         UpDown          =   -1  'True
         CurrentDate     =   38085
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tgl RencanaKeluar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kondisi Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   11160
         TabIndex        =   29
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cara Pulang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9120
         TabIndex        =   28
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Status Keluar "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   27
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tujuan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kelas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7320
         TabIndex        =   25
         Top             =   360
         Width           =   405
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   11880
      Picture         =   "frmRencanaPindahPulangPasien.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRencanaPindahPulangPasien.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmRencanaPindahPulangPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSimpan_Click()
On Error GoTo pesan
If Periksa("datacombo", dcStatusKeluar, "Status Keluar Kosong") = False Then Exit Sub
If dcKelas.Enabled = True Then
    If Periksa("datacombo", dcKelas, "Kelas Pelayanan Masih Kosong") = False Then Exit Sub
End If
If sp_StrukOrder = False Then Exit Sub
If sp_rencanaPindahPulangPasien("A") = False Then Exit Sub

MsgBox "Data Berhasil Di Simpan", vbInformation, "Informasi"
cmdSimpan.Enabled = False
Exit Sub
pesan:
Call msubPesanError
End Sub

Private Function sp_StrukOrder() As Boolean
    On Error GoTo Errload
    sp_StrukOrder = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglRencanaKeluar.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText)
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
        '.Parameters.Append .CreateParameter("NoOrderGudang", adChar, adParamInput, 20, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoOrder", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk order", vbCritical, "Validasi"
            sp_StrukOrder = False
        Else
            txtNoOrder.Text = .Parameters("OutputNoOrder").value
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
Errload:
    Call msubPesanError(" sp_StrukOrder")
    sp_StrukOrder = False
'    Resume 0
End Function


Private Function sp_rencanaPindahPulangPasien(f_status As String) As Boolean
On Error GoTo Errload
    sp_rencanaPindahPulangPasien = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, mstrNoPakai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtnopendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtnocm.Text)
        .Parameters.Append .CreateParameter("TglRencanaKeluar", adDate, adParamInput, , Format(dtpTglRencanaKeluar.value, "yyyy-MM-dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdStatusKeluar", adChar, adParamInput, 2, IIf(dcStatusKeluar.BoundText = "", Null, dcStatusKeluar.BoundText))
        .Parameters.Append .CreateParameter("KdStatusPulang", adChar, adParamInput, 2, IIf(dcStatusPulang.BoundText = "", Null, dcStatusPulang.BoundText))
        .Parameters.Append .CreateParameter("KdKondisiPulang", adChar, adParamInput, 2, IIf(dcKondisiPulang.BoundText = "", Null, dcKondisiPulang.BoundText))
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, IIf(dcKelas.BoundText = "", Null, dcKelas.BoundText))
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, IIf(dcRuanganTujuan.BoundText = "", Null, dcRuanganTujuan.BoundText))
        .Parameters.Append .CreateParameter("NamaTempatTujuan", adChar, adParamInput, 150, IIf(dcRuanganTujuan.Text = "", Null, dcRuanganTujuan.Text))
        .Parameters.Append .CreateParameter("TglKeluar", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_RencanaPindahPulangPasien"
        .CommandType = adCmdStoredProc

        .Execute

        If Not (.Parameters("RETURN_VALUE").value = 0) Then
            sp_rencanaPindahPulangPasien = False
            MsgBox "Ada kesalahan penyimpanan data", vbCritical, "Validasi"

        End If
        Set dbcmd = Nothing
    End With
    Exit Function
Errload:
    Call msubPesanError("sp_rencanaPindahPulangPasien")
'    Resume 0
End Function

Private Sub cmdTutup_Click()
Unload Me
End Sub



Private Sub dcKelas_GotFocus()
'    select distinct * from V_KelasPelayanan
'    If KeyAscii = 39 Then KeyAscii = 0
'    If KeyAscii = 13 Then
'        If dcKondisiPulang.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "select  distinct Kdkelas,Kelas from V_KelasPelayanan where kdinstalasi='" & mstrKdInstalasiLogin & "' and KdRuangan = '" & dcRuanganTujuan.BoundText & "' and Expr1='1' and StatusEnabled='1' and Expr2='1' and Expr3='1' "
        Call msubDcSource(dcKelas, rs, strSQL)
'        If rs.EOF = True Then dcKelas.Text = "": Exit Sub
'        dcKelas.BoundText = rs(0).Value
'        dcKelas.Text = rs(1).Value

'     End If
    
End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "select distinct Kdkelas,Kelas from V_KelasPelayanan where kdinstalasi='" & mstrKdInstalasiLogin & "' and KdRuangan ='" & dcRuanganTujuan.BoundText & "' and Kelas like '%" & dcKelas.Text & "%' and Expr1='1' and StatusEnabled='1' and Expr2='1' and Expr3='1'  "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcKelas.Text = "": Exit Sub
        dcKelas.BoundText = rs(0).value
        dcKelas.Text = rs(1).value

     End If
    
End Sub

Private Sub dcKondisiPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcKondisiPulang.MatchedWithList = True Then cmdSimpan.SetFocus
        strSQL = "select KdKondisiPulang, KondisiPulang from KondisiPulang where KdKondisiPulang in ('01','06','07') and StatusEnabled = 1 and KondisiPulang like '%" & dcKondisiPulang.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcKondisiPulang.Text = "": Exit Sub
        dcKondisiPulang.BoundText = rs(0).value
        dcKondisiPulang.Text = rs(1).value

     End If

End Sub

Private Sub dcKondisiPulang_LostFocus()
    If dcKondisiPulang.MatchedWithList = False Then dcKondisiPulang.Text = ""
End Sub

Private Sub dcRuanganTujuan_Change()
    dcKelas.Text = ""
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then
  If dcRuanganTujuan.MatchedWithList = True Then dcKelas.SetFocus
        strSQL = "SELECT KdRuangan,NamaRuangan FROM Ruangan WHERE (NamaRuangan LIKE '%" & dcRuanganTujuan.Text & "%') and Kdinstalasi ='03' order by NamaRuangan "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuanganTujuan.BoundText = rs(0).value
        dcRuanganTujuan.Text = rs(1).value
End If
    
End Sub


Private Sub dcStatusKeluar_Change()
On Error GoTo pesan

'strSQL = "SELECT KdStatusKeluar, StatusKeluar FROM StatusKeluarKamar WHERE kdStatusKeluar ='" & dcStatusKeluar.BoundText & "'"
'Call msubRecFO(rs, strSQL)
'If rs.EOF Then Exit Sub

If dcStatusKeluar.BoundText = "02" Then
    dcStatusPulang.Enabled = True
    dcStatusPulang.SetFocus
    dcKondisiPulang.Enabled = True

    dcRuanganTujuan.Text = ""
'    dcKelas.Text = ""
    dcRuanganTujuan.Enabled = False
    dcKelas.Enabled = False

Else
    dcRuanganTujuan.Enabled = True
    dcRuanganTujuan.SetFocus
    dcKelas.Enabled = True

    dcStatusPulang.Text = ""
    dcKondisiPulang.Text = ""
    dcStatusPulang.Enabled = False
    dcKondisiPulang.Enabled = False


End If
Exit Sub
pesan:
Call msubPesanError
End Sub

'Private Sub dcStatusKeluar_Click(Area As Integer)
''On Error GoTo pesan
'
''strSQL = "SELECT KdStatusKeluar, StatusKeluar FROM StatusKeluarKamar WHERE kdStatusKeluar ='" & dcStatusKeluar.BoundText & "'"
''Call msubRecFO(rs, strSQL)
''If rs.EOF Then Exit Sub
''
''If rs(0).Value = "02" Then
''    dcStatusPulang.Enabled = True
''    dcStatusPulang.SetFocus
''    dcKondisiPulang.Enabled = True
''
''    dcRuanganTujuan.Text = ""
'''    dcKelas.Text = ""
''    dcRuanganTujuan.Enabled = False
''    dcKelas.Enabled = False
''
''Else
''    dcRuanganTujuan.Enabled = True
''    dcRuanganTujuan.SetFocus
''    dcKelas.Enabled = True
''
''    dcStatusPulang.Text = ""
''    dcKondisiPulang.Text = ""
''    dcStatusPulang.Enabled = False
''    dcKondisiPulang.Enabled = False
''
''
''End If
''Exit Sub
''pesan:
''Call msubPesanError
'End Sub





'Private Sub dcStatusKeluar_GotFocus()
''On Error GoTo pesan
''
''strSQL = "SELECT KdStatusKeluar, StatusKeluar FROM StatusKeluarKamar WHERE kdStatusKeluar ='" & dcStatusKeluar.BoundText & "'"
''Call msubRecFO(rs, strSQL)
''If rs.EOF Then Exit Sub
''
''If rs(0).Value = "02" Then
''    dcStatusPulang.Enabled = True
''    dcStatusPulang.SetFocus
''    dcKondisiPulang.Enabled = True
''
''    DcRuanganTujuan.Text = ""
'''    dcKelas.Text = ""
''    DcRuanganTujuan.Enabled = False
''    dcKelas.Enabled = False
''
''Else
''    DcRuanganTujuan.Enabled = True
''    DcRuanganTujuan.SetFocus
''    dcKelas.Enabled = True
''
''    dcStatusPulang.Text = ""
''    dcKondisiPulang.Text = ""
''    dcStatusPulang.Enabled = False
''    dcKondisiPulang.Enabled = False
''
''
''End If
''Exit Sub
''pesan:
''Call msubPesanError
'End Sub

Private Sub dcStatusKeluar_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0

If KeyAscii = 13 Then
  If dcStatusKeluar.MatchedWithList = True Then
    If dcStatusPulang.Enabled = True Then
        dcStatusPulang.SetFocus
    Else
        dcRuanganTujuan.SetFocus
    End If
    Exit Sub
                
        strSQL = "SELECT KdStatusKeluar, StatusKeluar FROM StatusKeluarKamar  WHERE (StatusKeluar LIKE '%" & dcStatusKeluar.Text & "%') "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcStatusKeluar.BoundText = rs(0).value
        dcStatusKeluar.Text = rs(1).value
    End If
End If
    
End Sub

Private Sub dcStatusPulang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If dcStatusPulang.MatchedWithList = True Then dcKondisiPulang.SetFocus
        strSQL = "select KdStatusPulang, StatusPulang from StatusPulang where StatusEnabled =1 and StatusPulang like '%" & dcStatusPulang.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then dcStatusPulang.Text = "": Exit Sub
        dcStatusPulang.BoundText = rs(0).value
        dcStatusPulang.Text = rs(1).value

     End If
End Sub

Private Sub dcStatusPulang_LostFocus()
    If dcStatusKeluar.MatchedWithList = False Then dcStatusKeluar.Text = ""
End Sub

Private Sub dtpTglRencanaKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcStatusKeluar.SetFocus
End Sub

Private Sub Form_Load()
    
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpTglRencanaKeluar.value = Now
    Call subLoadDcSource

    dcStatusPulang.Enabled = False
    dcKondisiPulang.Enabled = False
    dcKelas.Enabled = False
    dcRuanganTujuan.Enabled = False
    
End Sub


Private Sub subLoadDcSource()

'strSQL = "select KdKelas,DeskKelas from KelasPelayanan where StatusEnabled = 1"
strSQL = "select distinct Kdkelas,Kelas from V_KelasPelayanan where kdinstalasi='" & mstrKdInstalasiLogin & "' and Expr1='1' and StatusEnabled='1' and Expr2='1' and Expr3='1'"
Call msubDcSource(dcKelas, rs, strSQL)

strSQL = "(select KdRuangan,NamaRuangan from Ruangan where kdInstalasi='03' and StatusEnabled = 1)"
Call msubDcSource(dcRuanganTujuan, rs, strSQL)

strSQL = "select KdKondisiPulang, KondisiPulang from KondisiPulang where KdKondisiPulang in ('01','06','07') and StatusEnabled = 1"
Call msubDcSource(dcKondisiPulang, rs, strSQL)

strSQL = "select KdStatusPulang, StatusPulang from StatusPulang where StatusEnabled =1"
Call msubDcSource(dcStatusPulang, rs, strSQL)

strSQL = "select KdStatusKeluar, StatusKeluar from StatusKeluarKamar where kdStatusKeluar in ('01','02') and StatusEnabled=1 "
Call msubDcSource(dcStatusKeluar, rs, strSQL)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDaftarPasienRI.Enabled = True

End Sub
