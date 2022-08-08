VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPesanDiet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Pesan Menu Pasien"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesanMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   11025
   Begin MSDataGridLib.DataGrid dgUserPemesan 
      Height          =   2055
      Left            =   3480
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
   Begin VB.Frame fraPasien 
      Caption         =   "Daftar Pasien"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   23
      Top             =   960
      Width           =   11055
      Begin VB.CheckBox chkPilihSemua 
         Caption         =   "Pilih Semua"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkPasien 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         MaskColor       =   &H80000005&
         TabIndex        =   25
         Top             =   480
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid fgPasien 
         Height          =   3255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         HighLight       =   0
         Appearance      =   0
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
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   0
      TabIndex        =   20
      Top             =   6480
      Width           =   11055
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   9960
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   9960
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid fgDiet 
         Height          =   2295
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         Appearance      =   0
      End
   End
   Begin VB.TextBox txtNoOrder 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesan Menu Diet"
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
      TabIndex        =   0
      Top             =   4920
      Width           =   11055
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   9960
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   9960
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtJumlahOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6240
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtUserPemesan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8160
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkUserPemesan 
         Caption         =   "Pemesan"
         Height          =   255
         Left            =   8160
         TabIndex        =   2
         Top             =   840
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dcJenisMenuDiet 
         Height          =   330
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSDataListLib.DataCombo dcJenisWaktu 
         Height          =   330
         Left            =   8160
         TabIndex        =   6
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSComCtl2.DTPicker dtpTglOrder 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   125960195
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin MSDataListLib.DataCombo dcKategoriDiet 
         Height          =   330
         Left            =   6240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSDataListLib.DataCombo DcKeterangan 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Waktu"
         Height          =   210
         Left            =   8160
         TabIndex        =   18
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah Menu"
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kategori Diet"
         Height          =   210
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Pesan"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diet"
         Height          =   210
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
   Begin VB.Label lblPasien 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   26
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9360
      Picture         =   "frmPesanMenu.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPesanMenu.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmPesanDiet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subSetGrid()
    On Error Resume Next
    With fgDiet
        .Cols = 12
        .Rows = 2

        .ColWidth(0) = 2000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(1) = 1500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
        .ColWidth(4) = 1300
        .ColAlignment(4) = flexAlignCenterCenter
        .ColWidth(5) = 2000
        .ColAlignment(5) = flexAlignCenterCenter
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 1000
        .ColAlignment(11) = flexAlignRightCenter

        .TextMatrix(0, 0) = "Nama Pasien"
        .TextMatrix(0, 1) = "Jenis Diet"
        .TextMatrix(0, 2) = "Kategori Diet"
        .TextMatrix(0, 3) = "KdJenisWaktu"
        .TextMatrix(0, 4) = "Waktu Diet"
        .TextMatrix(0, 5) = "Keterangan"
        .TextMatrix(0, 6) = "No Order"
        .TextMatrix(0, 7) = "KdJenisMenuDiet"
        .TextMatrix(0, 8) = "KdKategoryDiet"
        .TextMatrix(0, 9) = "KdKeterangan"
        .TextMatrix(0, 10) = "No Pakai"
        .TextMatrix(0, 11) = "Jml Order"
    End With
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad
    strSQL = ""
    strSQL = "select KdJenisMenuDiet, JenisMenuDiet from JenisMenuDiet_V order by JenisMenuDiet"
    Call msubDcSource(dcJenisMenuDiet, rs, strSQL)

    strSQL = "select KdJenisWaktu, JenisWaktu from JenisWaktu order by JenisWaktu"
    Call msubDcSource(DcJenisWaktu, rs, strSQL)

    strSQL = "select KdKategoryDiet, KategoryDiet from KategoryDiet order by KategoryDiet"
    Call msubDcSource(dcKategoriDiet, rs, strSQL)

    strSQL = "select KdKeterangan, Keterangan from KeteranganMenuDiet order by Keterangan"
    Call msubDcSource(DcKeterangan, rs, strSQL)
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub ChkPasien_Click()
    On Error GoTo errLoad
    If ChkPasien.Value = vbChecked Then
        fgPasien.TextMatrix(fgPasien.Row, 0) = Chr$(187)
        fgPasien.TextMatrix(fgPasien.Row, 10) = 1
    Else
        fgPasien.TextMatrix(fgPasien.Row, 0) = ""
        fgPasien.TextMatrix(fgPasien.Row, 10) = 0
    End If
    Exit Sub
errLoad:
    msubPesanError
End Sub

Private Sub chkPasien_LostFocus()
    ChkPasien.Visible = False
End Sub

Private Sub chkPilihSemua_Click()
    If chkPilihSemua.Value = Checked Then
        With fgPasien
            For i = 1 To fgPasien.Rows - 1
                ChkPasien.Value = vbChecked
                .TextMatrix(i, 0) = Chr$(187)
                .TextMatrix(i, 10) = 1
            Next i
        End With
    Else
        With fgPasien
            For i = 1 To fgPasien.Rows - 1
                ChkPasien.Value = vbUnchecked
                .TextMatrix(i, 0) = ""
                .TextMatrix(i, 10) = 0
            Next i
        End With
    End If

End Sub

Private Sub chkUserPemesan_Click()
    On Error GoTo errLoad
    If chkUserPemesan.Value = 0 Then
        txtUserPemesan.Enabled = False
        txtUserPemesan.Text = ""
        If dgUserPemesan.Visible = True Then dgUserPemesan.Visible = False
    Else
        txtUserPemesan.Enabled = True
        strSQL = "SELECT IdPegawai, NamaLengkap FROM V_DataPegawai"
        Call msubRecFO(rs, strSQL)
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkUserPemesan.Value = 0 Then
            dcJenisdiet.SetFocus
        Else
            txtUserPemesan.SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
    TxtNoOrder.Text = ""
    dtpTglPesan = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""
    dcJenisMenuDiet.Text = ""
    dcKategoriDiet.Text = ""
    DcJenisWaktu.Text = ""
    DcKeterangan.Text = ""
    txtJumlahOrder.Text = ""
    mstrIdPegawai = ""
    fgDiet.clear
    Call subSetGrid
    dgUserPemesan.Visible = False
    CmdSimpan.Enabled = True
End Sub

Private Sub cmdSimpan_Click()
    With fgDiet
    If fgDiet.TextMatrix(1, 0) = "" Then MsgBox "Menu yang akan dipesan masih kosong", vbExclamation, "Validasi": Exit Sub
        For i = 1 To .Rows - 2
            If SimpanStrukOrder() = False Then Exit Sub
            If SimpanDetailOrderDietPasien(TxtNoOrder.Text, .TextMatrix(i, 10), .TextMatrix(i, 3), .TextMatrix(i, 7), .TextMatrix(i, 9), .TextMatrix(i, 8), .TextMatrix(i, 11), "A") = False Then Exit Sub
            TxtNoOrder.Text = ""
        Next i
    End With
    MsgBox "menu sudah Tersimpan", vbInformation, "Informasi"
    fgDiet.clear
    Call subSetGrid
    CmdSimpan.Enabled = False
    cmdTutup.SetFocus
End Sub

Private Sub cmdTambah_Click()
    On Error GoTo errLoad
    strSQL = ""
    If Periksa("datacombo", dcJenisMenuDiet, "Jenis diet kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKategoriDiet, "Detail diet kosong") = False Then Exit Sub
    If Periksa("datacombo", DcJenisWaktu, "Waktu diet kosong") = False Then Exit Sub
    If Periksa("text", txtJumlahOrder, "Jumlah Order kosong") = False Then Exit Sub
    If txtJumlahOrder <= 0 Then MsgBox "Jumlah Order Tidak Boleh 0 ( Nol ) atau Kurang dari 0 ( Nol )", vbExclamation, "Informasi": txtJumlahOrder.SetFocus: Exit Sub
    For i = 0 To fgDiet.Rows - 1
        If dcJenisMenuDiet.BoundText = fgDiet.TextMatrix(i, 7) And DcJenisWaktu.BoundText = fgDiet.TextMatrix(i, 3) Then
            'MsgBox "Pasien" & " " & fgDiet.TextMatrix(i, 0) & " " & "sudah memesan menu untuk waktu" & " " & dcJenisWaktu.Text, vbExclamation, "Validasi"
            MsgBox "Menu tersebut sudah dientry", vbExclamation, "Informasi"
            dcJenisdiet.SetFocus
            Exit Sub
        End If
    Next i

    For i = 0 To fgDiet.Rows - 2
        With fgPasien
            For n = 1 To .Rows - 1
                If .TextMatrix(n, 10) = "1" Then
                    If .TextMatrix(n, 3) = fgDiet.TextMatrix(i, 0) And fgDiet.TextMatrix(i, 3) = DcJenisWaktu.BoundText Then
                        MsgBox "Pasien tidak bisa memesan 2 menu untuk waktu yang sama, ulangi pemesanan!!!", vbExclamation, "Validasi"
                        Exit Sub
                    End If
                End If
            Next n
        End With
    Next i

    strSQL = "SELECT TglOrder, NoOrder, NoPendaftaran, [Nama Pasien]," & _
    "KdRuangan, KdJenisMenuDiet, KdKategoryDiet, KdJenisWaktu, NoPakai,NamaRuangan " & _
    "From dbo.PesanMenuDiet_V " & _
    "Where KdRuangan Like '%" & mstrKdRuangan & "%' AND Year(TglOrder)='" & Year(dtpTglOrder.Value) & "' AND mONTH(TglOrder)='" & Month(dtpTglOrder.Value) & "'  AND day(TglOrder)='" & Day(dtpTglOrder.Value) & "' AND NoKirim IS NULL"
    Call msubRecFO(rs, strSQL)

    With fgPasien
        For m = 1 To .Rows - 1
            If .TextMatrix(m, 10) = "1" Then
                For l = 1 To rs.RecordCount
                    If .TextMatrix(m, 3) = rs(3).Value And DcJenisWaktu.BoundText = rs(7).Value And .TextMatrix(m, 8) = rs(8).Value And Format(dtpTglOrder.Value, "dd/mm/yyyy") = Format(rs(0).Value, "dd/mm/yyyy") Then
                        MsgBox "Pasien" & " " & .TextMatrix(m, 3) & " " & " sudah memesan menu untuk waktu" & " " & DcJenisWaktu.Text & ", " & "Lihat daftar pesanan pasien di Ctrl+Z !!!" & " " & "Silahkan ulangi pemesanan", vbExclamation, "Informasi"
                        Exit Sub
                    End If
                    rs.MoveNext
                Next l
            End If
            strSQL = "SELECT TglOrder, NoOrder, NoPendaftaran, [Nama Pasien]," & _
            "KdRuangan, KdJenisMenuDiet, KdKategoryDiet, KdJenisWaktu, NoPakai,NamaRuangan " & _
            "From dbo.PesanMenuDiet_V " & _
            "Where KdRuangan Like '%" & mstrKdRuangan & "%' AND Year(TglOrder)='" & Year(dtpTglOrder.Value) & "' AND mONTH(TglOrder)='" & Month(dtpTglOrder.Value) & "'  AND day(TglOrder)='" & Day(dtpTglOrder.Value) & "' AND NoKirim IS NULL"
            Call msubRecFO(rs, strSQL)
        Next m
    End With
    For k = 1 To fgPasien.Rows - 1
        If fgPasien.TextMatrix(k, 10) = "1" Then
            With fgDiet
                .TextMatrix(.Rows - 1, 0) = fgPasien.TextMatrix(k, 3)
                .TextMatrix(.Rows - 1, 1) = dcJenisMenuDiet.Text
                .TextMatrix(.Rows - 1, 2) = dcKategoriDiet.Text
                .TextMatrix(.Rows - 1, 3) = DcJenisWaktu.BoundText
                .TextMatrix(.Rows - 1, 4) = DcJenisWaktu.Text
                If DcKeterangan.Text = "" Then
                    .TextMatrix(.Rows - 1, 5) = "-"
                    .TextMatrix(.Rows - 1, 9) = "22"
                Else
                    .TextMatrix(.Rows - 1, 5) = DcKeterangan.Text
                    .TextMatrix(.Rows - 1, 9) = DcKeterangan.BoundText
                End If
                .TextMatrix(.Rows - 1, 6) = TxtNoOrder.Text
                .TextMatrix(.Rows - 1, 7) = dcJenisMenuDiet.BoundText
                .TextMatrix(.Rows - 1, 8) = dcKategoriDiet.BoundText
                .TextMatrix(.Rows - 1, 10) = fgPasien.TextMatrix(k, 8)
                .TextMatrix(.Rows - 1, 11) = txtJumlahOrder.Text
                .Rows = .Rows + 1
            End With
        End If
    Next k
    strSQL = ""
    TxtNoOrder.Text = ""
    dtpTglOrder = Now
    chkUserPemesan.Value = Checked
    txtUserPemesan.Text = ""
    dcJenisMenuDiet.Text = ""
    dcKategoriDiet.Text = ""
    DcKeterangan.BoundText = ""
    DcJenisWaktu.Text = ""
    txtJumlahOrder.Text = ""
    CmdSimpan.Enabled = True
    dgUserPemesan.Visible = False
    dtpTglOrder.SetFocus
    Exit Sub
errLoad:
    '    Call msubPesanError
End Sub

Private Sub cmdTambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdSimpan.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
    If strKet = "1" Then
        frmDaftarPasienRI.Enabled = True
        Unload frmDaftarPasienPesanMenuGizi
    Else
        frmDaftarPasienPesanMenuGizi.Enabled = True
        Unload frmDaftarPasienRI
    End If
End Sub

Private Sub dcDetailDiet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DcKeterangan.SetFocus
End Sub

Private Sub dcJenisMenuDiet_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJenisMenuDiet.Text)) = 0 Then dcKategoriDiet.SetFocus: Exit Sub
        If dcJenisMenuDiet.MatchedWithList = True Then dcKategoriDiet.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdJenisMenuDiet, JenisMenuDiet from JenisMenuDiet WHERE JenisMenuDiet LIKE '%" & dcJenisMenuDiet.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcJenisMenuDiet.BoundText = dbRst(0).Value
        dcJenisMenuDiet.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisWaktu_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(DcJenisWaktu.Text)) = 0 Then DcKeterangan.SetFocus: Exit Sub
        If DcJenisWaktu.MatchedWithList = True Then DcKeterangan.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdJenisWaktu, JenisWaktu from JenisWaktu WHERE JenisWaktu LIKE '%" & DcJenisWaktu.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        DcJenisWaktu.BoundText = dbRst(0).Value
        DcJenisWaktu.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcKategoriDiet_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcKategoriDiet.Text)) = 0 Then DcJenisWaktu.SetFocus: Exit Sub
        If dcKategoriDiet.MatchedWithList = True Then DcJenisWaktu.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdKategoryDiet, KategoryDiet from KategoryDiet WHERE KategoryDiet LIKE '%" & dcKategoriDiet.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcKategoriDiet.BoundText = dbRst(0).Value
        dcKategoriDiet.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub DcKeterangan_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(DcKeterangan.Text)) = 0 Then txtJumlahOrder.SetFocus: Exit Sub
        If DcKeterangan.MatchedWithList = True Then txtJumlahOrder.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "select KdKeterangan, Keterangan from KeteranganMenuDiet WHERE Keterangan LIKE '%" & DcKeterangan.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        DcKeterangan.BoundText = dbRst(0).Value
        DcKeterangan.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgUserPemesan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgUserPemesan
    WheelHook.WheelHook dgUserPemesan
End Sub

Private Sub dgUserPemesan_DblClick()
    Call dgUserPemesan_KeyPress(13)
End Sub

Private Sub dgUserPemesan_KeyPress(KeyAscii As Integer)
    If dgUserPemesan.ApproxCount = 0 Then Exit Sub
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        txtUserPemesan.Text = dgUserPemesan.Columns(1).Value
        mstrIdPegawai = dgUserPemesan.Columns(0).Value
        If mstrIdPegawai = "" Then
            MsgBox "Pilih dulu User Pemesan yang menangani Pasien", vbCritical, "Validasi"
            txtUserPemesan.Text = ""
            dgUserPemesan.SetFocus
            Exit Sub
        End If
        chkUserPemesan.Value = 1
        dgUserPemesan.Visible = False
        cmdTambah.SetFocus
    End If
    If KeyAscii = 27 Then
        dgUserPemesan.Visible = False
    End If
End Sub

Private Sub dtpTglOrder_Change()
    dtpTglOrder.MaxDate = Now
End Sub

Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisMenuDiet.SetFocus
End Sub

Private Sub fgPasien_Click()
    On Error GoTo hell
    'If chkPasien.Value = vbChecked Then Exit Sub
    If fgPasien.Rows = 1 Then Exit Sub
    If fgPasien.Col <> 0 Then Exit Sub
    ChkPasien.Visible = True
    ChkPasien.Top = fgPasien.RowPos(fgPasien.Row) + 250
    Dim intChk As Integer
    intChk = ((fgPasien.ColPos(fgPasien.Col + 1) - fgPasien.ColPos(fgPasien.Col)) / 2)
    ChkPasien.Left = fgPasien.ColPos(fgPasien.Col) + intChk  ' - 250  '+ intChk
    ChkPasien.SetFocus
    If fgPasien.Col <> 0 Then
        If fgPasien.TextMatrix(fgPasien.Row, 0) <> "" Then
            ChkPasien.Value = 1
        Else
            ChkPasien.Value = 0
        End If
    End If
    Exit Sub
hell:
    Call msubPesanError

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call openConnection
    dtpTglOrder.Value = Now
    dgUserPemesan.Left = 2400
    dgUserPemesan.Top = 3000
    Call subSetGrid
    Call subLoadDcSource
    Call cmdBatal_Click
    Call load_dataPasien
    lblPasien.Caption = fgPasien.Rows - 1 & " " & "Pasien"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If strKet = "1" Then
        frmDaftarPasienRI.Enabled = True
        Unload frmDaftarPasienPesanMenuGizi
    Else
        frmDaftarPasienPesanMenuGizi.Enabled = True
        Unload frmDaftarPasienRI
    End If
End Sub

Private Sub txtJmlPesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlPesan.SetFocus
End Sub

Private Sub txtJumlahOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdTambah.SetFocus
    If KeyAscii > 48 And KeyAscii > 58 Then KeyAscii = 0
End Sub

Private Sub txtUserPemesan_Change()
    Call subLoadUserPemesan
End Sub

'untuk meload data User Pemesan di grid
Private Sub subLoadUserPemesan()
    strSQL = "SELECT IdPegawai AS [Id Pegawai], NamaLengkap AS [Nama User Pemesan], JenisPegawai FROM V_DataPegawai     mstrFilterUserPemesan WHERE NamaLengkap like '%" & txtUserPemesan.Text & "%'"
    dgUserPemesan.Visible = True

    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgUserPemesan
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 3000
    End With
    dgUserPemesan.Left = 3480
    dgUserPemesan.Top = 6360
End Sub

Private Sub txtUserPemesan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUserPemesan.Text = "" Then
            MsgBox "Isi dulu User Pemesannya.", vbExclamation, "Validasi"
            txtUserPemesan.SetFocus
        Else
           dgUserPemesan.Visible = True
            dgUserPemesan.SetFocus
        End If
    ElseIf KeyAscii = 27 Then
        dgUserPemesan.Visible = False
    End If
End Sub

Private Function SimpanStrukOrder() As Boolean
    '====================================
    'simpan Struk Order
    '====================================
    SimpanStrukOrder = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, TxtNoOrder.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtpTglOrder.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, "131")
        .Parameters.Append .CreateParameter("KdSupplier", adChar, adParamInput, 4, Null)
        'add by JDR (2009-07-22)
        If mstrIdPegawai = "" Then
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        Else
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, mstrIdPegawai)
        End If
        .Parameters.Append .CreateParameter("OutKode", adChar, adParamOutput, 10, Null)
        .ActiveConnection = dbConn
        .CommandText = "Add_StrukOrder"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data struk order", vbCritical, "Validasi"
            sp_StrukOrder = False
        Else
            TxtNoOrder.Text = .Parameters("OutKode").Value
        End If
    End With
End Function

Private Function SimpanPesanMenuDiet() As Boolean
    SimpanPesanMenuDiet = True
    '====================================
    'simpan Pesan menu Diet
    '====================================
    Dim i As Integer
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, TxtNoOrder)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, strNoPakai)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("NoCM", adVarChar, adParamInput, 12, mstrNoCM)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

        .ActiveConnection = dbConn
        .CommandText = "Add_PesanMenuDietPasien"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("return_value").Value = 0) Then
            SimpanPesanMenuDiet = False
            MsgBox "Ada kesalahan dalam pemasukan data Detail Pesan Menu Diet", vbExclamation, "Validasi"
        Else
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Sub txtUserPemesan_LostFocus()
    'dgUserPemesan.Visible = False
End Sub

Private Function SimpanDetailOrderDietPasien(F_NoOrder As String, f_NoPakai As String, F_KdJenisWaktu As String, F_KdJenisMenuDiet As String, f_Keterangan As String, F_KdKategoryDiet As String, F_JmlOrder As String, f_status As String) As Boolean
    SimpanDetailOrderDietPasien = True
    '================================
    'Simpan Detail Order Menu Diet
    '================================
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, F_NoOrder)
        .Parameters.Append .CreateParameter("KdJenisWaktu", adChar, adParamInput, 3, F_KdJenisWaktu)
        .Parameters.Append .CreateParameter("KdJenisMenuDiet", adChar, adParamInput, 3, F_KdJenisMenuDiet)
        .Parameters.Append .CreateParameter("NoPakai", adChar, adParamInput, 10, f_NoPakai)
        .Parameters.Append .CreateParameter("KdKeterangan", adChar, adParamInput, 2, f_Keterangan)
        .Parameters.Append .CreateParameter("KdKategoryDiet", adVarChar, adParamInput, 3, F_KdKategoryDiet)
        .Parameters.Append .CreateParameter("JmlOrder", adTinyInt, adParamInput, , CInt(F_JmlOrder))
        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
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

Private Sub subloadGridpasien()
    With fgPasien
        .Rows = 2
        .Cols = 11
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "No Pendaftaran"
        .TextMatrix(0, 2) = "NoCM"
        .TextMatrix(0, 3) = "Nama Pasien"
        .TextMatrix(0, 4) = "JK"
        .TextMatrix(0, 5) = "Umur"
        .TextMatrix(0, 6) = "Kelas"
        .TextMatrix(0, 7) = "Jenis Pasien"
        .TextMatrix(0, 8) = "No Pakai"
        .TextMatrix(0, 9) = "Tgl Pendafataran"
        .TextMatrix(0, 10) = ""
    End With
End Sub

Private Sub subClearGridPasien()
    With fgPasien
        .Rows = 2
        .Cols = 11
        .ColWidth(0) = 300
        .ColWidth(1) = 1500
        .ColAlignment(1) = flexAlignCenterCenter
        .ColWidth(2) = 800
        .ColAlignment(2) = flexAlignCenterCenter
        .ColWidth(3) = 2100
        .ColWidth(4) = 300
        .ColAlignment(4) = flexAlignCenterCenter
        .ColWidth(5) = 1500
        .ColWidth(6) = 1200
        .ColWidth(7) = 1500
        .ColAlignment(7) = flexAlignLeftCenter
        .ColWidth(8) = 1200
        .ColAlignment(8) = flexAlignLeftCenter
        .ColWidth(9) = 2200
        .ColAlignment(9) = flexAlignLeftCenter
        .ColWidth(10) = 0
        Call subloadGridpasien
    End With
End Sub

Private Sub load_dataPasien()
    Dim i As Integer
    strSQL = ""
    strSQL = "select NoPendaftaran,NoCM,[Nama Pasien],JK,Umur,Kelas,JenisPasien,NoPakai,TglMasuk from V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        Call subClearGridPasien
        fgPasien.Rows = rs.RecordCount + 1
        For i = 1 To rs.RecordCount
            With fgPasien
                .TextMatrix(i, 0) = Chr$(187)
                '.TextMatrix(i, 0) = ""
                .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(0).Value), "-", rs.Fields(0)) 'No Pendaftaran
                .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(1).Value), "-", rs.Fields(1)) 'NoCM
                .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(2).Value), "-", rs.Fields(2)) 'Nama Pasien
                .TextMatrix(i, 4) = IIf(IsNull(rs.Fields(3).Value), "-", rs.Fields(3)) 'JK
                .TextMatrix(i, 5) = IIf(IsNull(rs.Fields(0).Value), "-", rs.Fields(4)) 'Kelas
                .TextMatrix(i, 6) = IIf(IsNull(rs.Fields(1).Value), "-", rs.Fields(5)) 'Umur
                .TextMatrix(i, 7) = IIf(IsNull(rs.Fields(2).Value), "-", rs.Fields(6)) 'Jenis Pasien
                .TextMatrix(i, 8) = IIf(IsNull(rs.Fields(3).Value), "-", rs.Fields(7)) 'No Pakai
                .TextMatrix(i, 9) = IIf(IsNull(rs.Fields(1).Value), "-", rs.Fields(8)) 'Tgl Pendafataran
                .TextMatrix(i, 10) = 1
            End With
            rs.MoveNext
        Next i
    Else
        Call subClearGridPasien
    End If

End Sub

