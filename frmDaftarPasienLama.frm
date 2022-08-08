VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDaftarPasienLama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Daftar Pasien Lama"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDaftarPasienLama.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14670
   Begin VB.Frame frameJudul 
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
      TabIndex        =   8
      Top             =   1080
      Width           =   14655
      Begin VB.OptionButton optSemua 
         Caption         =   "&Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optPulang 
         Caption         =   "Pulan&g"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optPindahan 
         Caption         =   "&Pindahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
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
         Left            =   7800
         TabIndex        =   9
         Top             =   120
         Width           =   6735
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   345
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   124190723
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   345
            Left            =   3960
            TabIndex        =   1
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   124190723
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3600
            TabIndex        =   10
            Top             =   300
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgPasienLama 
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9128
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
               LCID            =   1033
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
               LCID            =   1033
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
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   7440
      Width           =   14655
      Begin VB.CommandButton cmdTP 
         Caption         =   "&Transaksi Pelayanan"
         Height          =   450
         Left            =   8880
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdBatalKeluar 
         Caption         =   "&Batal Keluar Kamar"
         Height          =   450
         Left            =   10920
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   440
         Width           =   2655
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   450
         Left            =   12960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Nama Pasien / No. CM"
         Height          =   210
         Left            =   1560
         TabIndex        =   7
         Top             =   200
         Width           =   2640
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8280
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5186
            Text            =   "Rincian Biaya Pelayanan (F1)"
            TextSave        =   "Rincian Biaya Pelayanan (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5293
            Text            =   "Rincian Kumulatif Biaya (Ctrl+F1)"
            TextSave        =   "Rincian Kumulatif Biaya (Ctrl+F1)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5186
            Text            =   "Cetak Daftar Pasien Lama (F11)"
            TextSave        =   "Cetak Daftar Pasien Lama (F11)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4833
            MinWidth        =   2187
            Text            =   "Refresh Data (F5)"
            TextSave        =   "Refresh Data (F5)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5186
            Text            =   "Ubah Kamar (Ctrl+F7)"
            TextSave        =   "Ubah Kamar (Ctrl+F7)"
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
      Left            =   12840
      Picture         =   "frmDaftarPasienLama.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDaftarPasienLama.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDaftarPasienLama.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmDaftarPasienLama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'untuk load data pasien di form ubah kamar dan bed pasien
Private Sub subLoadFormUbahKamarBed()
    On Error GoTo hell
    mstrNoPen = dgPasienLama.Columns(0).Value
    mstrNoCM = dgPasienLama.Columns(1).Value
    mstrKdKelas = dgPasienLama.Columns("KdKelas").Value
    With frmUbahKamardanBed
        .Show
        .txtNoPakai.Text = dgPasienLama.Columns("NoPakai").Value
        .txtNoPendaftaran.Text = dgPasienLama.Columns(0).Value
        .txtNoCM.Text = mstrNoCM
        .txtNamaPasien.Text = dgPasienLama.Columns(2).Value
        If dgPasienLama.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtThn.Text = dgPasienLama.Columns(10).Value
        .txtBln.Text = dgPasienLama.Columns(11).Value
        .txtHari.Text = dgPasienLama.Columns(12).Value

        .txtKdRuanganAsal.Text = dgPasienLama.Columns(16).Value
        .txtRuangPerawatan.Text = mstrNamaRuangan
        .dcKelasPK.BoundText = mstrKdKelas
        .txtNoKamLama.Text = dgPasienLama.Columns(8)
        .txtNoBedLama.Text = dgPasienLama.Columns("NoBed")
    End With
    Exit Sub
hell:
End Sub

Private Sub cmdBatalKeluar_Click()
    On Error GoTo errLoad

    If dgPasienLama.ApproxCount = 0 Then Exit Sub
    Set rs = Nothing
'    strSQL = "SELECT NoPendaftaran FROM BiayaPelayanan WHERE NoPendaftaran ='" & dgPasienLama.Columns("No. Registrasi") & "' AND KdRuangan ='" & dgPasienLama.Columns("KdRuangan") & "' AND NoStruk is null"
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = True Then
'        MsgBox "Pasien " & dgPasienLama.Columns("Nama Pasien").Value & " Sudah Bayar ! ", vbCritical, "Validasi"
'        Exit Sub
'    End If

    '*******************
    ''add splakuk 2009-02-05 utk validasi error pasien pindah² lebih dari 2x
    If dgPasienLama.Columns("Cara Keluar").Value = "Pindah Kamar" Then
        If Len(Trim(dgPasienLama.Columns("TglPulang"))) <> 0 Then
            MsgBox "Pasien " & dgPasienLama.Columns("Nama Pasien").Value & " Sudah dipulangkan dari Ruang Rawat Inap pindahan! ", vbCritical, "Validasi"
            Exit Sub
        End If
    End If

    If dgPasienLama.Columns("Cara Keluar").Value = "Pindah Kamar" Then
        strSQL = "SELECT * FROM PindahKamarPerawatanRI WHERE NoPendaftaran ='" & dgPasienLama.Columns("No. Registrasi") & "' AND KdRuangan <>'" & dgPasienLama.Columns("KdRuangan") & "' AND StatusMasukKamar='T'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            strSQL = "SELECT * FROM PasienPulang WHERE NoPendaftaran ='" & dgPasienLama.Columns("No. Registrasi") & "' "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            If rs.EOF = True Then
                MsgBox "Pasien " & dgPasienLama.Columns("Nama Pasien").Value & " sudah di pindahkan lagi ke Ruang Rawat Inap lain, dan belum dimasukkan di ruangan tersebut, cek Ruangan terakhir/Ruang lain! ", vbCritical, "Validasi"
                Exit Sub
            End If
        End If
    End If
    ''********************

    If dgPasienLama.ApproxCount = 0 Then Exit Sub
    If MsgBox("Yakin akan membatalkan KELUAR KAMAR " & vbNewLine & "pasien " & dgPasienLama.Columns("Nama Pasien").Value & "", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, dgPasienLama.Columns("NoPakai").Value)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgPasienLama.Columns("No. Registrasi").Value)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, dgPasienLama.Columns("NoCM").Value)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglKeluar", adDate, adParamInput, , IIf(Len(Trim(dgPasienLama.Columns("TglKeluar"))) = 0, Null, Format(dgPasienLama.Columns("TglKeluar"), "yyyy/MM/dd HH:mm:ss")))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 2, IIf(Len(Trim(dgPasienLama.Columns("TglPulang"))) = 0, "PK", "PU"))
        .Parameters.Append .CreateParameter("OutputMsg", adChar, adParamOutput, 1, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PasienBatalKeluarKamar"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam pembatalan pasien", vbCritical, "Validasi"
        Else
            If LCase(.Parameters("OutputMsg").Value) = "t" Then
                MsgBox "Pasien " & dgPasienLama.Columns("Nama Pasien").Value & " sudah masuk ruangan " & dgPasienLama.Columns("RuanganTujuan") & "", vbCritical, "Validasi"
            Else
                MsgBox "Pasien " & dgPasienLama.Columns("Nama Pasien").Value & " Batal keluar kamar", vbInformation, "Informasi"
            End If
            Call Add_HistoryLoginActivity("Add_PasienBatalKeluarKamar")
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
    Set dbcmd = Nothing
    Call cmdCari_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdCari_Click()
    On Error GoTo hell

    Set rs = Nothing
    strSQLX = ""
    FlagStatusPulang = ""

    If optPindahan.Value = True Then
        FlagStatusPulang = "1"
        If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
            rs.Open "select top 100 * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select top 100 * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        Else
            rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is null and RuanganTujuan is not null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        End If
    End If

    If optPulang.Value = True Then
        FlagStatusPulang = "2"
        If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
            rs.Open "select top 100 * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select top 100 * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        Else
            rs.Open "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select * from V_DaftarPasienLamaRI where tglPulang is not null and RuanganTujuan is null and ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        End If
    End If

    If optsemua.Value = True Then
        FlagStatusPulang = "3"
        If dtpAwal.Day <> dtpAkhir.Day Or dtpAwal.Month <> dtpAkhir.Month Or dtpAwal.Year <> dtpAkhir.Year Then
            rs.Open "select top 100 * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select top 100 * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        Else
            rs.Open "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'", dbConn, adOpenStatic, adLockOptimistic
            strSQLX = "select * from V_DaftarPasienLamaRI where ([NoCM] like '%" & txtParameter.Text & "%' OR [Nama Pasien] like '%" & txtParameter.Text & "%' OR Alamat like '%" & txtParameter.Text & "%') AND (TglKeluar between '" & Format(dtpAwal.Value, "yyyy/MM/dd HH:mm:00") & "' and '" & Format(dtpAkhir.Value, "yyyy/MM/dd HH:mm:59") & "') and KdRuangan='" & mstrKdRuangan & "'"
        End If
    End If

    Set dgPasienLama.DataSource = rs
    
    'Pasien umum yg sudah keluar tidak bisa di ubah data transaksinya (503)
'    If dgPasienLama.Columns("JenisPasien") = "UMUM" Then
'        cmdTP.Enabled = False
'    Else
'        cmdTP.Enabled = True
'    End If
    
    
    Call SetGridPasienLamaRI
    If dgPasienLama.ApproxCount = 0 Then dtpAwal.SetFocus Else dgPasienLama.SetFocus
    lblJumData.Caption = "Data 0/" & rs.RecordCount

    Exit Sub
hell:
End Sub

Private Sub cmdTP_Click()
    On Error GoTo hell
    If dgPasienLama.Columns(0).Value = "" Then
        Exit Sub
    End If

'    If sp_UpdateJmlPelayananKamarBK(dgPasienLama.Columns("No. Registrasi").Value) = False Then Exit Sub
    Call subLoadFormTP
    Exit Sub
hell:
    MsgBox "Silahkan hubungi Administrator untuk input sewa kamar", vbCritical, "Perhatian"
    Call subLoadFormTP
End Sub

Private Sub cmdtutup_Click()
    Unload Me
End Sub

Private Sub dgPasienLama_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPasienLama
    WheelHook.WheelHook dgPasienLama
End Sub

Private Sub dgPasienLama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBatalKeluar.SetFocus
End Sub

Private Sub dgPasienLama_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    
    'Pasien umum yg sudah keluar tidak bisa di ubah data transaksinya (503)
'    If dgPasienLama.Columns("JenisPasien") = "UMUM" Then
'    cmdTP.Enabled = False
'    Else
'    cmdTP.Enabled = True
'    End If
    
    lblJumData.Caption = "Data " & dgPasienLama.Bookmark & "/" & dgPasienLama.ApproxCount
End Sub

Private Sub dtpAkhir_Change()
    On Error Resume Next
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpAwal_Change()
    On Error Resume Next
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Public Sub PostingHutangPenjaminPasien_AU(strStatus As String)
    On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, dgPasienLama.Columns("No. Registrasi").Value)
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
            If dgPasienLama.ApproxCount = 0 Then Exit Sub
            mstrTglKeluar = dgPasienLama.Columns("TglKeluar")
            mstrNoPen = dgPasienLama.Columns("No. Registrasi").Value
            If Me.dgPasienLama.Columns("JenisPasien") <> "UMUM" Then Call PostingHutangPenjaminPasien_AU("A")
            mstrFormPengirim = Me.Name
            strCetak = IIf(strCtrlKey = "2", "Lengkap", "Singkat") ' Cek Ctrl+F1

            frm_cetak_RincianBiaya.Show

        Case vbKeyF5
            Call cmdCari_Click
        Case vbKeyF11
            If dgPasienLama.ApproxCount = 0 Then Exit Sub
            mdTglAwal = dtpAwal.Value: mdTglAkhir = dtpAkhir.Value
            frmCetakDaftarPasienLama.Show

        Case vbKeyF7
            If strCtrlKey = 4 Then
                'hanya Petugas SIMRS saja
                If boolStafSIMRS = True Then
                    If dgPasienLama.ApproxCount = 0 Then Exit Sub
                    frmDaftarPasienLama.Enabled = False
                    mstrKdSubInstalasi = dgPasienLama.Columns("KdSubInstalasi").Value
                    mstrKdKelas = dgPasienLama.Columns("KdKelas").Value
                    Call subLoadFormUbahKamarBed
                Else
                    MsgBox "Otoritas transaksi ini tidak berlaku bagi Anda." & vbCrLf & "Hubungi SIMRS", vbOKOnly + vbCritical
                End If
            End If
    End Select
    Exit Sub
errLoad:

End Sub

Private Sub Form_Load()
    On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    StatusBar1.Panels(5).Visible = False
    dtpAwal.Value = Format(Now, "dd MMM yyyy 00:00:00")
    dtpAkhir.Value = Now
    frameJudul.Caption = "Daftar Pasien Lama " + mstrNamaRuangan
    
    Call cmdCari_Click
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub SetGridPasienLamaRI()
    With dgPasienLama
        .Columns(0).Width = 1200
        .Columns(0).Alignment = dbgCenter
        .Columns(0).Caption = "No. Registrasi"
        .Columns(1).Width = 1300
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1800
        .Columns(3).Width = 300
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 4400
        .Columns(5).Width = 1590
        .Columns(6).Width = 1590
        .Columns(7).Width = 1500
        .Columns(8).Width = 1800
        .Columns(9).Width = 1590
        .Columns(10).Width = 1500
        .Columns(11).Width = 1500
        .Columns(12).Width = 1200
        .Columns(13).Width = 1300
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
    End With
End Sub

'untuk load data pasien lama di form transaksi pelayanan
Private Sub subLoadFormTP()
    On Error GoTo hell

    mstrNoPen = dgPasienLama.Columns("No. Registrasi").Value
    mstrNoCM = dgPasienLama.Columns("NoCM").Value

    With frmTransaksiPasien
        .Show
        .txtNoPendaftaran.Text = dgPasienLama.Columns("No. Registrasi").Value
        .txtNoCM.Text = dgPasienLama.Columns("NoCM").Value
        .txtNamaPasien.Text = dgPasienLama.Columns("Nama Pasien").Value
        If dgPasienLama.Columns(3).Value = "P" Then
            .txtSex.Text = "Perempuan"
        Else
            .txtSex.Text = "Laki-Laki"
        End If
        .txtKls.Text = dgPasienLama.Columns("Kelas").Value
        .txtThn.Text = dgPasienLama.Columns("Thn").Value
        .txtBln.Text = dgPasienLama.Columns("Bln").Value
        .txtHr.Text = dgPasienLama.Columns("Hr").Value

        .txtJenisPasien.Text = dgPasienLama.Columns("JenisPasien").Value
        .txtTglDaftar.Text = dgPasienLama.Columns("TglPendaftaran").Value

        'check in
        If mblnAdmin = True Then
            .cmdTambahPT.Enabled = True
            .cmdHapusDataPT.Enabled = True
            .cmdUbahPT.Enabled = True
        Else
            .cmdTambahPT.Enabled = False
            .cmdHapusDataPT.Enabled = False
            .cmdUbahPT.Enabled = False
        End If

        '*******
        mdTglMasuk = dgPasienLama.Columns("TglMasuk").Value
        mstrKdKelas = dgPasienLama.Columns("KdKelas").Value
        mstrKdSubInstalasi = dgPasienLama.Columns("KdSubInstalasi").Value
    End With

    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub optPindahan_Click()
    Call cmdCari_Click
End Sub

Private Sub optPulang_Click()
    Call cmdCari_Click
End Sub

Private Sub optSemua_Click()
    Call cmdCari_Click
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
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

