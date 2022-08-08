VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmVerifikasiData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifist2000 - Verifikasi Data Pasien"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerifikasiData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14805
   Begin VB.Frame fraCetakUlangStrukBKM 
      Height          =   3135
      Left            =   2640
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton cmdTutupUlang 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   8760
         TabIndex        =   18
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetakUlang 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgCetakUlangStrukBKMVerifikasi 
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   10335
         _ExtentX        =   18230
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
   End
   Begin MSDataGridLib.DataGrid dgDaftarPasien 
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8916
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   14775
      Begin VB.Frame Frame4 
         Caption         =   "Kategori Periode"
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
         Left            =   5760
         TabIndex        =   7
         Top             =   120
         Width           =   8895
         Begin VB.OptionButton optTglKeluar 
            Caption         =   "TglKeluar"
            Height          =   375
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optTglMasuk 
            Caption         =   "TglMasuk"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   8040
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   3120
            TabIndex        =   19
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   53739523
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   5760
            TabIndex        =   20
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   53739523
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   5400
            TabIndex        =   8
            Top             =   315
            Width           =   255
         End
      End
      Begin VB.Label lblJumData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data 0/0"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   7680
      Width           =   14775
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   7800
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdUbahKelPasien 
         Caption         =   "&Jenis Pasien"
         Height          =   495
         Left            =   9495
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   12885
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdTagihan 
         Caption         =   "&Tagihan Pasien"
         Height          =   495
         Left            =   11190
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   8565
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13018
            Text            =   "Rincian Biaya Sementara (F1)"
            TextSave        =   "Rincian Biaya Sementara (F1)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13018
            Text            =   "Cetak Kwitansi (Shift+F1)"
            TextSave        =   "Cetak Kwitansi (Shift+F1)"
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
      TabIndex        =   12
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
   Begin MSDataListLib.DataCombo dcInstalasiPelayanan 
      Height          =   330
      Left            =   360
      TabIndex        =   21
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcRuangPerawatan 
      Height          =   330
      Left            =   2400
      TabIndex        =   22
      Top             =   2280
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcJenisPasien 
      Height          =   330
      Left            =   4440
      TabIndex        =   23
      Top             =   2280
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcKelas 
      Height          =   330
      Left            =   5640
      TabIndex        =   24
      Top             =   2280
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcRuangPerawatanTerahkir 
      Height          =   330
      Left            =   6960
      TabIndex        =   25
      Top             =   2280
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcKondisiPulang 
      Height          =   330
      Left            =   8640
      TabIndex        =   26
      Top             =   2280
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmVerifikasiData.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   12960
      Picture         =   "frmVerifikasiData.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmVerifikasiData.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmVerifikasiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Dim rsB As New ADODB.recordset
Private Sub subLoadDcSource()
    Call msubDcSource(dcInstalasiPelayanan, rs, "SELECT KdInstalasi, NamaInstalasi FROM Instalasi")
    Call msubDcSource(dcRuangPerawatan, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan")
    Call msubDcSource(dcJenisPasien, rs, "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien")
    Call msubDcSource(dcKelas, rs, "SELECT     KdKelas, DeskKelas FROM KelasPelayanan")
    Call msubDcSource(dcRuangPerawatanTerahkir, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan")
    Call msubDcSource(dcKondisiPulang, rs, "SELECT KdKondisiPulang, KondisiPulang FROM KondisiPulang")
    
    
End Sub

Public Sub cmdCari_Click()
On Error Resume Next
    MousePointer = vbHourglass
    Call subLoadDataPasien
    MousePointer = vbDefault
End Sub

Private Sub cmdCetak_Click()
    If dgDaftarPasien.ApproxCount = 0 Then Exit Sub
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    frmCetakDaftarPasienVerifikasi.Show
End Sub

Private Sub cmdCetakUlang_Click()
    If dgCetakUlangStrukBKMVerifikasi.ApproxCount = 0 Then Exit Sub
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    frmCetakUlangStrukBKMVerifikasi.Show
End Sub

Private Sub cmdTagihan_Click()
On Error GoTo errLoad

    'cek pasien RI
    If dgDaftarPasien.ApproxCount = 0 Then Exit Sub
    
    strSQL = "SELECT * FROM RegistrasiIGD WHERE NoPendaftaran = '" & dgDaftarPasien.Columns("NoPendaftaran").Value & "' AND StatusPulang = 'T'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then
        MsgBox "Pasien belum keluar dari IGD", vbCritical
        Exit Sub
    End If
    
    strSQL = "SELECT NamaRuangan FROM v_PasienAktifPakaiKamar WHERE NoPendaftaran='" _
        & dgDaftarPasien.Columns("NoPendaftaran").Value & "'"
    Call msubRecFO(rs, strSQL)
    If rs.RecordCount <> 0 Then MsgBox "Pasien belum keluar dari Rawat Inap ( " & rs(0) & " )", vbCritical, "Validasi": dgDaftarPasien.SetFocus: Exit Sub
    
    strSQL = "SELECT KdKelompokPasien, IdPenjamin FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & dgDaftarPasien.Columns("NoPendaftaran").Value & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
    End If
    
    If mstrKdPenjaminPasien <> "2222222222" Then
        strSQL = "SELECT * FROM PemakaianAsuransi WHERE NoPendaftaran='" & dgDaftarPasien.Columns("NoPendaftaran").Value & "'"
        Call msubRecFO(rsB, strSQL)
        If rsB.RecordCount = 0 Then
            MsgBox "Lengkapi dahulu data penjamin pasien", vbCritical, "Validasi"
            Call cmdUbahKelPasien_Click
            Exit Sub
        End If
    End If
    
    
    
    Call subLoadFormTP
    

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTutupUlang_Click()
       fraCetakUlangStrukBKM.Visible = False
End Sub

Private Sub cmdUbahKelPasien_Click()
On Error GoTo hell
    Call subLoadFormJP
hell:
End Sub

Private Sub dcInstalasiPelayanan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcInstalasiPelayanan.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdInstalasi, NamaInstalasi FROM Instalasi where NamaInstalasi LIKE '%" & dcInstalasiPelayanan.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcInstalasiPelayanan.BoundText = rs(0).Value
        dcInstalasiPelayanan.Text = rs(1).Value
    End If
End Sub

Private Sub dcJenisPasien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcJenisPasien.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKelompokPasien, JenisPasien FROM KelompokPasien where JenisPasien LIKE '%" & dcJenisPasien.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJenisPasien.BoundText = rs(0).Value
        dcJenisPasien.Text = rs(1).Value
    End If

End Sub

Private Sub dcKelas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcKelas.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT     KdKelas, DeskKelas FROM KelasPelayanan where DeskKelas LIKE '%" & dcKelas.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKelas.BoundText = rs(0).Value
        dcKelas.Text = rs(1).Value
    End If

End Sub

Private Sub dcKondisiPulang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcKondisiPulang.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdKondisiPulang, KondisiPulang FROM KondisiPulang where KondisiPulang LIKE '%" & dcKondisiPulang.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcKondisiPulang.BoundText = rs(0).Value
        dcKondisiPulang.Text = rs(1).Value
    End If

End Sub

Private Sub dcRuangPerawatan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcRuangPerawatan.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan where NamaRuangan LIKE '%" & dcRuangPerawatan.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangPerawatan.BoundText = rs(0).Value
        dcRuangPerawatan.Text = rs(1).Value
    End If

End Sub

Private Sub dcRuangPerawatanTerahkir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If dcRuangPerawatanTerahkir.MatchedWithList = True Then cmdCari.SetFocus
        strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan where NamaRuangan LIKE '%" & dcRuangPerawatanTerahkir.Text & "%' "
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcRuangPerawatanTerahkir.BoundText = rs(0).Value
        dcRuangPerawatanTerahkir.Text = rs(1).Value
    End If

End Sub

Private Sub dgDaftarPasien_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 0
            mstrFilter = " ORDER BY InstalasiPelayanan"
        Case 1
            mstrFilter = " ORDER BY RuanganPerawatan"
        Case 2
            mstrFilter = " ORDER BY NoPendaftaran"
        Case 3
            mstrFilter = " ORDER BY NamaPasien"
        Case 4
            mstrFilter = " ORDER BY JK"
        Case 5
            mstrFilter = " ORDER BY Umur"
        Case 6
            mstrFilter = " ORDER BY TglPendaftaran"
        Case 7
            mstrFilter = " ORDER BY JenisPasien"
        Case 8
            mstrFilter = " ORDER BY Kelas"
        Case 9
            mstrFilter = " ORDER BY TglKeluar"
        Case 10
            mstrFilter = " ORDER BY KondisiPulang"
        Case 11
            mstrFilter = " ORDER BY TglPendaftaran"
        Case Else
            mstrFilter = ""
    End Select
    Call cmdCari_Click

End Sub

Private Sub dgDaftarPasien_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
        lblJumData.Caption = "Data " & dgDaftarPasien.Bookmark & "/" & dgDaftarPasien.ApproxCount
    
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
    cmdCari_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errLoad
    Dim strShiftKey As String
    strShiftKey = (Shift + vbShiftMask)
    Select Case KeyCode
        Case vbKeyF1
        If dgDaftarPasien.ApproxCount = 0 Then Exit Sub
            If strShiftKey = 2 Then
                mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").Value
                Call subLoadCetakUlangStrukBKMVerifikasi
                    fraCetakUlangStrukBKM.Visible = True
            Else
                mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").Value
                frm_cetak_RincianBiaya.Show
            End If
    End Select

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    dtpAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
    optTglMasuk.Value = True
    'blnFrmCariPasien = True
    mstrFilter = ""
    Call subLoadDataPasien
    Call subLoadDcSource
End Sub

'untuk load data pasien di form transaksi pelayanan
Private Sub subLoadFormTP()
    On Error GoTo hell
    mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").Value
    'mstrNoCM = dgPasien.Columns(1).Value
    With frmTagihanPasienVerifikasi
        .Show
        .txtnopendaftaran.Text = mstrNoPen
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasien.Columns("NamaPasien").Value
        .txtSex.Text = dgDaftarPasien.Columns("JK").Value
        .txtThn.Text = dgDaftarPasien.Columns("UmurTahun")
        .txtBln.Text = dgDaftarPasien.Columns("UmurBulan")
        .txtHari.Text = dgDaftarPasien.Columns("UmurHari")
        .txtJenisPasien.Text = dgDaftarPasien.Columns("JenisPasien").Value
        Call .txtNoPendaftaran_KeyPress(13)
    End With
hell:
End Sub

'untuk load data pasien
Private Sub subLoadDataPasien()
On erro GoTo errLoad
    If optTglMasuk.Value = True Then
        strSQL = "SELECT InstalasiPelayanan, RuanganPerawatan,JK, UmurTahun,UmurBulan, UmurHari, JenisPasien, Kelas, RuanganPerawatanTerakhir, KondisiPulang,TglMasuk,NamaPasien,NoPendaftaran" & _
                 " FROM V_DaftarPasienAllforVerifikasi" & _
        " WHERE TglMasuk BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND InstalasiPelayanan LIKE '%" & dcInstalasiPelayanan.Text & "%' AND RuanganPerawatan LIKE '%" & dcRuangPerawatan.Text & "%' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND KdKelas like '%" & dcKelas.BoundText & "%' AND RuanganPerawatanTerakhir LIKE '%" & dcRuangPerawatanTerahkir.Text & "%' AND KondisiPulang LIKE '%" & dcKondisiPulang.Text & "%'" & mstrFilter
    ElseIf optTglKeluar.Value = True Then
        strSQL = "SELECT InstalasiPelayanan, RuanganPerawatan,JK, UmurTahun,UmurBulan,UmurHari, JenisPasien, Kelas, RuanganPerawatanTerakhir, KondisiPulang,TglMasuk,NamaPasien,NoPendaftaran" & _
                 " FROM V_DaftarPasienAllforVerifikasi" & _
        " WHERE TglKeluar BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'  " & _
         " AND InstalasiPelayanan LIKE '%" & dcInstalasiPelayanan.Text & "%' AND RuanganPerawatan LIKE '%" & dcRuangPerawatan.Text & "%' AND JenisPasien LIKE '%" & dcJenisPasien.Text & "%' AND KdKelas like '%" & dcKelas.BoundText & "%' AND RuanganPerawatanTerakhir LIKE '%" & dcRuangPerawatanTerahkir.Text & "%' AND KondisiPulang LIKE '%" & dcKondisiPulang.Text & "%'" & mstrFilter
    End If
    
    Call msubRecFO(rs, strSQL)
    Set dgDaftarPasien.DataSource = rs
    With dgDaftarPasien
        .Columns("InstalasiPelayanan").Width = 2200
        .Columns("RuanganPerawatan").Width = 2000
        .Columns("JenisPasien").Width = 1100
        .Columns("Kelas").Width = 1100
        .Columns("KondisiPulang").Width = 1200
        .Columns("NamaPasien").Width = 2500
        .Columns("NoPendaftaran").Width = 0
        .Columns("JK").Width = 0
        .Columns("UmurTahun").Width = 0
        .Columns("UmurBulan").Width = 0
        .Columns("UmurHari").Width = 0
        If optTglMasuk.Value = True Then
            .Columns("TglMasuk").Width = 2000
        ElseIf optTglKeluar.Value = True Then
            .Columns("TglKeluar").Width = 2000
        End If
        
        
        
        
    End With
Exit Sub
errLoad:
    msubPesanError
End Sub
'untuk load data Struk BKM verifikasi
Private Sub subLoadCetakUlangStrukBKMVerifikasi()
        strSQL = "SELECT NoPendaftaran, NoStruk, NoBKM, TglBKM, RuanganKasir, UserKasir" & _
                " FROM V_CetakUlangStrukBKMforVerifikasi" & _
                " WHERE NoPendaftaran = '" & mstrNoPen & "' "
    Call msubRecFO(rs, strSQL)
    Set dgCetakUlangStrukBKMVerifikasi.DataSource = rs
    With dgCetakUlangStrukBKMVerifikasi
        .Columns("NoPendaftaran").Width = 1300
        .Columns("NoStruk").Width = 1300
        .Columns("NoBKM").Width = 1300
        .Columns("TglBKM").Width = 2000
        .Columns("RuanganKasir").Width = 2000
        .Columns("UserKasir").Width = 2100
    End With
End Sub



Private Sub Form_Unload(Cancel As Integer)
    blnFrmCariPasien = False
End Sub

Private Sub subLoadFormJP()
On Error GoTo hell
    mstrNoPen = dgDaftarPasien.Columns("NoPendaftaran").Value
    strSQL = "SELECT KdKelompokPasien, IdPenjamin, NoCM FROM V_KelasTanggunganPenjamin WHERE (NoPendaftaran = '" & mstrNoPen & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        mstrKdJenisPasien = rs("KdKelompokPasien").Value
        mstrKdPenjaminPasien = IIf(IsNull(rs("IdPenjamin")), "2222222222", rs("IdPenjamin"))
        mstrNoCM = rs("NoCM").Value
    End If

    With frmUbahJenisPasien
        .Show
        .txtNamaFormPengirim.Text = Me.Name
        Call msubRecFO(rs, "SELECT KdInstalasi FROM Ruangan WHERE (NamaRuangan = '" & dgDaftarPasien.Columns("RuanganPerawatan") & "')")
        .txtKdInstalasi.Text = rs("KdInstalasi")
        .txtnocm.Text = mstrNoCM
        .txtNamaPasien.Text = dgDaftarPasien.Columns("NamaPasien").Value
        If dgDaftarPasien.Columns("JK").Value = "P" Then
            .txtJK.Text = "Perempuan"
        Else
            .txtJK.Text = "Laki-laki"
        End If
        .txtThn.Text = dgDaftarPasien.Columns("UmurTahun")
        .txtBln.Text = dgDaftarPasien.Columns("UmurBulan")
        .txtHr.Text = dgDaftarPasien.Columns("UmurHari")
        .lblNoPendaftaran.Visible = False
        .txtnopendaftaran.Visible = False
'        .txtTglPendaftaran.Text = dgDaftarPasien.Columns("TglPendaftaran").Value
        .dcJenisPasien.BoundText = mstrKdJenisPasien
        .dcPenjamin.BoundText = mstrKdPenjaminPasien
    End With
hell:
End Sub

