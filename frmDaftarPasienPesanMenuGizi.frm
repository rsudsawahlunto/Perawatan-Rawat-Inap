VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDaftarPasienPesanMenuGizi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirrst 2000 - Daftar Pasien Pesan Menu Gizi"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienPesanMenuGizi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   14490
   Begin VB.Frame fraDaftar 
      Caption         =   "Daftar Pasien Pesan Menu Gizi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   14415
      Begin VB.Frame Frame1 
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
         Left            =   8415
         TabIndex        =   9
         Top             =   240
         Width           =   5775
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
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   11
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   146145283
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   12
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   146145283
            UpDown          =   -1  'True
            CurrentDate     =   38212
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   13
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.CheckBox ChkJenisWaktu 
         Caption         =   "Jenis Waktu"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox ChkPasien 
         Caption         =   "Pasien sudah Dikirim"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkJenis 
         Caption         =   "Jenis Diet"
         Height          =   240
         Left            =   5160
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcJenisdiet 
         Height          =   360
         Left            =   5160
         TabIndex        =   5
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataGridLib.DataGrid dgDaftarAntrianPasien 
         Height          =   4575
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   19
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
            Size            =   9.75
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
      Begin MSDataListLib.DataCombo DcJenisWaktu 
         Height          =   360
         Left            =   3120
         TabIndex        =   15
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.Label lblJumData 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1695
      End
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
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   14415
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Ubah Menu Gizi"
         Height          =   450
         Left            =   8400
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal Pesan Menu"
         Height          =   450
         Left            =   10320
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   400
         Width           =   3735
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   450
         Left            =   12270
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukan Nama Pasien /  No.CM "
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   2790
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
      Left            =   12720
      Picture         =   "frmDaftarPasienPesanMenuGizi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienPesanMenuGizi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienPesanMenuGizi.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmDaftarPasienPesanMenuGizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkJenis_Click()
    If chkJenis.Value = Checked Then
        dcJenisDiet.Enabled = True
        dcJenisDiet.SetFocus
    Else
        dcJenisDiet.Text = ""
        dcJenisDiet.Enabled = False
    End If
    Call cmdCari_Click
End Sub

Private Sub ChkJenisWaktu_Click()
    If ChkJenisWaktu.Value = Checked Then
        dcJenisWaktu.Enabled = True
        dcJenisWaktu.SetFocus
    Else
        dcJenisWaktu.Text = ""
        dcJenisWaktu.Enabled = False
    End If
    Call cmdCari_Click
End Sub

Private Sub ChkPasien_Click()
    Call cmdCari_Click
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo hell
    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    mstrNoOrder = dgDaftarAntrianPasien.Columns("NoOrder").Value
    mstrNoCM = dgDaftarAntrianPasien.Columns("NoCM")
    mstrKdJenisWaktu = dgDaftarAntrianPasien.Columns("KdJenisWaktu")
    mstrKdDiet = dgDaftarAntrianPasien.Columns("KdJenisMenuDiet")
    mstrKdKeterangan = dgDaftarAntrianPasien.Columns("KdKeterangan")
    mstrNoPakai = dgDaftarAntrianPasien.Columns("NoPakai")
    strJmlMenu = dgDaftarAntrianPasien.Columns("JmlOrder")
    If MsgBox("Yakin akan menghapus Daftar pesanan", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then Exit Sub

    strSQL = "delete From StrukOrder where NoOrder= '" & mstrNoOrder & "'"
    Call msubRecFO(rs, strSQL)
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"
    cmdCari_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Public Sub cmdCari_Click()
    Dim ruangandiet As String
    On Error GoTo hell
    LblJumData.Caption = ""
    MousePointer = vbHourglass
    If chkPasien.Value = vbUnchecked Then
        strSQL = "select * from PesanMenuDiet_V " & _
        " where KdJenisWaktu Like '%" & dcJenisWaktu.BoundText & "%' AND JenisMenuDiet Like '%" & dcJenisDiet.Text & "%'  AND ([Nama Pasien] like '%" & txtParameter.Text & "%'  OR NoCM like '%" & txtParameter.Text & "%') AND NoKirim Is Null AND (TglOrder between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' )AND KdRuangan Like '%" & mstrKdRuangan & "%'"
        Call msubRecFO(rs, strSQL)
        Set dgDaftarAntrianPasien.DataSource = rs
        cmdBatal.Enabled = True
        cmdEdit.Enabled = True
        Call SetGridPesanMenu
        LblJumData.Caption = dgDaftarAntrianPasien.ApproxCount & " Pemesanan "
        MousePointer = vbDefault
    Else
        strSQL = "select * from DaftarKirimMenuDiet_V " & _
        " where KdJenisWaktu Like '%" & dcJenisWaktu.BoundText & "%' AND JenisMenuDiet Like '%" & dcJenisDiet.Text & "%' AND([Nama Pasien] like '%" & txtParameter.Text & "%'  OR NoCM like '%" & txtParameter.Text & "%') AND (TglKirim between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "' )AND KdRuangan Like '%" & mstrKdRuangan & "%'"
        Call msubRecFO(rs, strSQL)
        Set dgDaftarAntrianPasien.DataSource = rs
        cmdBatal.Enabled = False
        cmdEdit.Enabled = False
        Call SetGridPasienGiziKirimMenuDiet
        LblJumData.Caption = dgDaftarAntrianPasien.ApproxCount & " Pengiriman "
        MousePointer = vbDefault

    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo hell
    If dgDaftarAntrianPasien.ApproxCount = 0 Then Exit Sub
    Call subloadFormPesanMenu
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisDiet_Change()
    Call cmdCari_Click
End Sub

Private Sub DcJenisWaktu_Click(Area As Integer)
    Call cmdCari_Click
End Sub

Private Sub dgPasienGizi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTP.SetFocus
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

Private Sub Form_Activate()
    txtParameter.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Set rs = Nothing
    rs.Open "select * from JenisWaktu", dbConn, adOpenStatic, adLockReadOnly
    Set dcJenisWaktu.RowSource = rs
    dcJenisWaktu.ListField = rs("JenisWaktu").Name
    dcJenisWaktu.BoundColumn = rs("KdJenisWaktu").Name
    Set rs = Nothing
    Set rs = Nothing
    rs.Open "select * from JenisMenuDiet", dbConn, adOpenStatic, adLockReadOnly
    Set dcJenisDiet.RowSource = rs
    dcJenisDiet.ListField = rs("JenisMenuDiet").Name
    dcJenisDiet.BoundColumn = rs("KdJenisMenuDiet").Name
    Set rs = Nothing
    dtpAkhir.Value = Now
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    Call cmdCari_Click
    ChkJenisWaktu.Value = Unchecked
    chkJenis.Value = vbUnchecked
    dcJenisWaktu.Enabled = False
    dcJenisDiet.Enabled = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPesanMenu()
    On Error Resume Next
    With dgDaftarAntrianPasien
        .Columns(0).Width = 1500
        .Columns(1).Width = 1200
        .Columns(2).Width = 1400
        .Columns(3).Width = 800
        .Columns(4).Width = 2000
        .Columns(5).Width = 300 'Jk
        .Columns(6).Width = 1500
        .Columns(7).Caption = "Waktu"
        .Columns(7).Width = 700
        .Columns(8).Caption = "Kelas"
        .Columns(8).Width = 1000
        .Columns(9).Width = 800
        .Columns(10).Width = 1000
        .Columns(11).Width = 1500
        .Columns(12).Width = 1200
        .Columns(13).Width = 1000
        .Columns(14).Width = 0
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
        .Columns(19).Width = 0
        .Columns(20).Width = 0
        .Columns(21).Width = 0
        .Columns(22).Width = 0
        .Columns(23).Width = 0
        .Columns(24).Width = 0
        .Columns(25).Width = 0
    End With
End Sub

Sub SetGridPasienGiziKirimMenuDiet()
    With dgDaftarAntrianPasien
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns(0).Width = 1100    'TglKirim
        .Columns(1).Width = 1000     'JenisWaktu

        .Columns(2).Width = 2500    '[Nama Pasien]
        .Columns(3).Width = 0     'JenisDiet
        .Columns(4).Width = 1500     'NamaDiet
        .Columns(5).Width = 2500    'Keterangan
        .Columns(6).Width = 0    'NoPendaftaran
        .Columns(7).Width = 800    'NoCM
        .Columns(8).Width = 300    'JK
        .Columns(9).Width = 500    'Berat badan/Tinggi badan
        .Columns(10).Width = 1200    'Umur
        .Columns(11).Width = 1200      'JenisPasien
        .Columns(12).Width = 975      'Kelas
        .Columns(13).Width = 0      'SubInstalasi
        .Columns(14).Width = 2000      'Ruangan
        .Columns(15).Width = 0      'TglMasuk
        .Columns(16).Width = 0      'NoKamar
        .Columns(17).Width = 0      'NoBed
        .Columns(18).Width = 0      'drPenangungJawab
        .Columns(19).Width = 0      'NoPakai
        .Columns(20).Width = 0      'thn
        .Columns(21).Width = 0      'bln
        .Columns(22).Width = 0      'hr
        .Columns(23).Width = 0      'kd jenistarif
        .Columns(24).Width = 0      'kdkelas
        .Columns(25).Width = 0      'kd subinstalasi
        .Columns(26).Width = 0      'kd kamar
        .Columns(27).Width = 0      'CaraMasuk
        .Columns(28).Width = 0      'StatusPulang
        .Columns(29).Width = 0      'TglKeluar
        .Columns(30).Width = 0      'StatusKeluar
        .Columns(31).Width = 0      'Alamat
        .Columns(32).Width = 1000      'TglPendaftaran
        .Columns(33).Width = 0   'kd dokter
        .Columns(34).Width = 0      ' Kd ruangan
        .Columns(35).Width = 1200      'NoKirim
        .Columns(36).Width = 0  'kd jns waktu
        .Columns(37).Width = 0  'kd jns diet
    End With
End Sub

Private Sub txtParameter_Change()
    On Error GoTo errLoad
    Call cmdCari_Click
    txtParameter.SetFocus
    txtParameter.SelStart = Len(txtParameter.Text)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subloadFormPesanMenu()
    mstrNoPen = dgDaftarAntrianPasien.Columns(2).Value
    mstrNoCM = dgDaftarAntrianPasien.Columns(3).Value
    mstrKelas = dgDaftarAntrianPasien.Columns(8).Value
    mstrNoOrder = dgDaftarAntrianPasien.Columns("NoOrder")
    mstrKdJenisWaktu = dgDaftarAntrianPasien.Columns("KdJenisWaktu")
    mstrKdDiet = dgDaftarAntrianPasien.Columns("KdJenisMenuDiet")
    mstrKdKeterangan = dgDaftarAntrianPasien.Columns("KdKeterangan")
    Me.Enabled = False
    strSQL = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and (NoPendaftaran like '%" & mstrNoPen & "%' or NoCM like '%" & mstrNoCM & "%') AND Kelas LIKE '%" & mstrKelas & "%'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Pasien sudah pulang", vbInformation + vbOKOnly, "Validasi": frmDaftarPasienPesanMenuGizi.Enabled = True: Exit Sub
    With frmEditMenuGizi
        .Show
        strSQL = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,TglMasuk,NoKamar,NoBed,NoPakai,UmurTahun,UmurBulan,UmurHari,KdSubInstalasi,KdKelas,CaraMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and (NoPendaftaran like '%" & mstrNoPen & "%' or NoCM like '%" & mstrNoCM & "%') AND Kelas LIKE '%" & mstrKelas & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then MsgBox "Pasien sudah pulang", vbInformation + vbOKOnly, "Validasi": frmDaftarPasienPesanMenuGizi.Enabled = True: Exit Sub
        .dcJenisDiet.SetFocus
        .txtNoPendaftaran.Text = rs(0).Value
        .txtNoCM.Text = rs(1).Value
        .txtNamaPasien.Text = rs(2).Value
        If rs(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtKls.Text = rs(5).Value
        .txtThn.Text = rs(11).Value
        .txtBln.Text = rs(12).Value
        .txtHr.Text = rs(13).Value
        .txtJenisPasien.Text = rs(6).Value
        .txtTglDaftar.Text = rs(7).Value
        mdTglMasuk = rs(7).Value
        mstrKdKelas = rs(5).Value
        strNoPakai = rs(10).Value
        mstrKdSubInstalasi = rs(14).Value
    End With
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Function SimpanDetailOrderDietPasien(F_NoOrder As String, f_NoPakai As String, F_KdJenisMenuDiet As String, F_KdJenisWaktu As String, F_KdKeterangan As String, F_JmlOrder As Integer, f_Status As String) As Boolean
    SimpanDetailOrderDietPasien = True
    '================================
    'Simpan Detail Order Menu Diet
    '================================
    Dim i As Integer
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, F_NoOrder)
        .Parameters.Append .CreateParameter("KdJenisWaktu", adChar, adParamInput, 3, F_KdJenisWaktu)
        .Parameters.Append .CreateParameter("KdJenisMenuDiet", adChar, adParamInput, 3, F_KdJenisMenuDiet)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, f_NoPakai)
        .Parameters.Append .CreateParameter("KdKeterangan", adChar, adParamInput, 2, F_KdKeterangan)
        .Parameters.Append .CreateParameter("KdKategoryDiet", adVarChar, adParamInput, 3, Null)
        .Parameters.Append .CreateParameter("JmlOrder", adTinyInt, adParamInput, , CInt(F_JmlOrder))
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailOrderJenisDietPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            SimpanDetailOrderDietPasien = False
            MsgBox "Ada kesalahan dalam pemasukan data Detail Struk Pesan", vbExclamation, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

