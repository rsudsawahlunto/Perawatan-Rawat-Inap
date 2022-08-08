VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPemesananDarah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pemesanan Darah"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPemesananDarah.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10725
   Begin VB.TextBox txtNoOrder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      MaxLength       =   10
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   465
      Left            =   7200
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   465
      Left            =   9000
      TabIndex        =   7
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   465
      Left            =   5400
      TabIndex        =   41
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detail Pemesanan Darah"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   37
      Top             =   4560
      Width           =   10695
      Begin MSDataListLib.DataCombo dcGolDarah 
         Height          =   330
         Left            =   2520
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcBentukDarah 
         Height          =   330
         Left            =   2520
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   4440
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "data Pesan Darah"
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
      TabIndex        =   27
      Top             =   2640
      Width           =   10695
      Begin VB.TextBox txtAlasan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1440
         Width           =   8415
      End
      Begin VB.TextBox txtNoHasilRad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtNoHasilLab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   32
         Top             =   720
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dcRuanganTujuan 
         Height          =   330
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcDokter 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo dcDiagnosa 
         Height          =   330
         Left            =   7560
         TabIndex        =   4
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSComCtl2.DTPicker dtpTglPesan 
         Height          =   330
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   496304131
         UpDown          =   -1  'True
         CurrentDate     =   38077
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Alasan Keperluan"
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   1500
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Hasil Radiologi"
         Height          =   210
         Index           =   3
         Left            =   5640
         TabIndex        =   35
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Hasil Laboratorium"
         Height          =   210
         Index           =   2
         Left            =   5640
         TabIndex        =   33
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruangan Tujuan"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Dokter Pemeriksa"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Diagnosa"
         Height          =   210
         Left            =   5640
         TabIndex        =   29
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tgl Pesan"
         Height          =   210
         Left            =   240
         TabIndex        =   28
         Top             =   420
         Width           =   795
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
      Height          =   1575
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   10695
      Begin VB.TextBox txtSubInstalasi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   17
         Top             =   1080
         Width           =   3015
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
         Height          =   315
         Left            =   7560
         MaxLength       =   6
         TabIndex        =   16
         Top             =   720
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
         Height          =   315
         Left            =   8340
         MaxLength       =   6
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtHr 
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
         Height          =   315
         Left            =   9120
         MaxLength       =   6
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7560
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNoCM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNoPendaftaran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNamaPasien 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Kasus Penyakit"
         Height          =   210
         Left            =   5640
         TabIndex        =   26
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "thn"
         Height          =   210
         Left            =   7995
         TabIndex        =   25
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "bln"
         Height          =   210
         Left            =   8790
         TabIndex        =   24
         Top             =   750
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "hr"
         Height          =   210
         Left            =   9570
         TabIndex        =   23
         Top             =   750
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Umur"
         Height          =   210
         Left            =   5640
         TabIndex        =   22
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   5640
         TabIndex        =   21
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. CM"
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No. Pendaftaran"
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pasien"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1020
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
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
      Left            =   8880
      Picture         =   "frmPemesananDarah.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPemesananDarah.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmPemesananDarah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdBatal_Click()
    Call subKosong
    Call subSetGrid

End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer
    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data darah harus diisi", vbExclamation, "Validasi": Exit Sub

    If sp_StrukOrder() = False Then Exit Sub
    If sp_DetailOrderPemakaianDarah("A") = False Then Exit Sub
    For i = 1 To fgData.Rows - 1
        If fgData.TextMatrix(i, 3) = "" Then GoTo keluar_
        With fgData
            If sp_DetailOrderDarah(.TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 2), "A") = False Then Exit Sub
        End With
    Next i
keluar_:
    MsgBox "No Order : " & txtNoOrder.Text, vbInformation, "Informasi"

    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_DetailOrderPemakaianDarah(f_status As String) As Boolean
    On Error GoTo errLoad
    sp_DetailOrderPemakaianDarah = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, txtNoPendaftaran.Text)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, txtNoCM.Text)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 6, Null)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(dcDokter.Text = "", Null, dcDokter.BoundText))
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, IIf(dcDiagnosa.Text = "", Null, dcDiagnosa.BoundText))
        .Parameters.Append .CreateParameter("AlasanKeperluan", adVarChar, adParamInput, 150, IIf(txtAlasan.Text = "", Null, txtAlasan.Text))
        .Parameters.Append .CreateParameter("NoHasilLabrujukan", adChar, adParamInput, 10, IIf(txtNoHasilLab.Text = "", Null, txtNoHasilLab.Text))
        .Parameters.Append .CreateParameter("NoHasilRadRujukan", adChar, adParamInput, 10, IIf(txtNoHasilRad.Text = "", Null, txtNoHasilRad.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailOrderPemakaianDarah"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data pemesanan darah", vbCritical, "Validasi"
            sp_DetailOrderPemakaianDarah = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError(" sp_DetailOrderPemakaianDarah")
End Function

Private Function sp_DetailOrderDarah(f_KdBentukDarah As String, f_KdGolDarah As String, f_JumlahBarang As Double, f_status As String) As Boolean
    On Error GoTo errLoad
    sp_DetailOrderDarah = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, txtNoOrder.Text)
        .Parameters.Append .CreateParameter("KdBentukDarah", adTinyInt, adParamInput, , f_KdBentukDarah)
        .Parameters.Append .CreateParameter("KdGolonganDarah", adChar, adParamInput, 2, f_KdGolDarah)
        .Parameters.Append .CreateParameter("JmlOrder", adDouble, adParamInput, , CDbl(f_JumlahBarang))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailOrderDarah"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data detail pemesanan darah", vbCritical, "Validasi"
            sp_DetailOrderDarah = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError(" AUD_DetailOrderDarah")
End Function

Private Function sp_StrukOrder() As Boolean
    On Error GoTo errLoad
    sp_StrukOrder = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglPesan.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, dcRuanganTujuan.BoundText)
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoOrder", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk order", vbCritical, "Validasi"
            sp_StrukOrder = False
        Else
            txtNoOrder.Text = .Parameters("OutputNoOrder").Value
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Exit Function
errLoad:
    Call msubPesanError(" sp_StrukOrder")
    sp_StrukOrder = False
End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Sub subKosong()
    On Error GoTo hell
    dtpTglPesan.Value = Now
    dcRuanganTujuan.Text = ""
    dcDokter.Text = ""
    txtAlasan.Text = ""
    dcDiagnosa.Text = ""
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subSetGrid()
    On Error GoTo hell
    With fgData
        .clear
        .Rows = 2
        .Cols = 5

        .RowHeight(0) = 400

        .TextMatrix(0, 0) = "Bentuk Darah"
        .TextMatrix(0, 1) = "Gol. Darah"
        .TextMatrix(0, 2) = "Qty Darah"
        .TextMatrix(0, 3) = "KdBentukDarah"
        .TextMatrix(0, 4) = "KdGolDarah"

        .ColWidth(0) = 5000
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 0
        .ColWidth(4) = 0

        .ColAlignment(2) = flexAlignRightCenter
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcDokter, rs, "Select IdPegawai,NamaLengkap From DataPegawai")
    Call msubDcSource(dcRuanganTujuan, rs, "Select KdRuangan,NamaRuangan From Ruangan Where KdRuangan<>'" & mstrKdRuangan & "' And StatusEnabled=1")
    Call msubDcSource(dcDiagnosa, rs, "Select Top 50 KdDiagnosa,NamaDiagnosa From Diagnosa Where StatusEnabled=1")
    Call msubDcSource(dcBentukDarah, rs, "Select KdBentukDarah,BentukDarah From BentukDarah Where StatusEnabled=1")
    Call msubDcSource(dcGolDarah, rs, "Select KdGolonganDarah,GolonganDarah From GolonganDarah Where StatusEnabled=1")
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left

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

Private Sub dcBentukDarah_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 0) = dcBentukDarah.Text
    fgData.TextMatrix(fgData.Row, 3) = dcBentukDarah.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcBentukDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcBentukDarah_Change
        dcBentukDarah.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcBentukDarah_LostFocus()
    dcBentukDarah.Visible = False
End Sub

Private Sub dcDiagnosa_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcDiagnosa.MatchedWithList = True Then fgData.SetFocus
        strSQL = "Select Top 50 KdDiagnosa,NamaDiagnosa From Diagnosa WHERE (NamaDiagnosa LIKE '" & dcDiagnosa.Text & "%') And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDiagnosa.BoundText = rs(0).Value
        dcDiagnosa.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcDokter_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcDokter.MatchedWithList = True Then txtAlasan.SetFocus
        strSQL = "Select IdPegawai,NamaLengkap From v_DataPegawai WHERE (NamaLengkap LIKE '" & dcDokter.Text & "%') AND KdJenisPegawai='001'"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDokter.BoundText = rs(0).Value
        dcDokter.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcGolDarah_Change()
    On Error GoTo errLoad
    fgData.TextMatrix(fgData.Row, 1) = dcGolDarah.Text
    fgData.TextMatrix(fgData.Row, 4) = dcGolDarah.BoundText
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcGolDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call dcGolDarah_Change
        dcGolDarah.Visible = False
        fgData.Col = 1
        fgData.SetFocus
    End If
End Sub

Private Sub dcGolDarah_LostFocus()
    dcGolDarah.Visible = False
End Sub

Private Sub dcRuanganTujuan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        If dcRuanganTujuan.MatchedWithList = True Then dcDokter.SetFocus
        strSQL = "Select KdRuangan,NamaRuangan From Ruangan WHERE (NamaRuangan LIKE '" & dcRuanganTujuan.Text & "%') AND KdRuangan<>'" & mstrKdRuangan & "' And StatusEnabled=1"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuanganTujuan.BoundText = rs(0).Value
        dcRuanganTujuan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dtpTglPesan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcRuanganTujuan.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then fgData.Rows = fgData.Rows + 1
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

            If fgData.Col > 7 Then
                fgData.Rows = fgData.Rows + 1
                fgData.Row = fgData.Rows - 1
                fgData.Col = 0
                fgData.SetFocus
            End If

        Case 27
            dgObatAlkes.Visible = False

        Case vbKeyDelete
            With fgData
                If .Row = .Rows Then Exit Sub
                If .Row = 0 Then Exit Sub

                If .Rows = 2 Then
                    For i = 0 To .Cols - 1
                        .TextMatrix(1, i) = ""
                    Next i
                    Exit Sub
                Else
                    .RemoveItem .Row
                End If
            End With

    End Select
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    Select Case fgData.Col
        Case 0 'bentuk dara
            Call subLoadDataCombo(dcBentukDarah)

        Case 1 'golonga  dara
            Call subLoadDataCombo(dcGolDarah)

        Case 2 'Jumlah
            txtIsi.MaxLength = 4
            Call subLoadText
            txtIsi.Text = Chr(KeyAscii)
            txtIsi.SelStart = Len(txtIsi.Text)

    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call subLoadDcSource
    Call cmdBatal_Click
End Sub

Private Sub txtAlasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcDiagnosa.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgData.Col
            Case 2

                fgData.TextMatrix(fgData.Row, 2) = txtIsi.Text
                fgData.Col = 0
                fgData.SetFocus
                SendKeys "{DOWN}"
                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
                    fgData.Rows = fgData.Rows + 1
                End If
        End Select

        txtIsi.Visible = False

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub
