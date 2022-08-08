VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmFilterJenisPeriksa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFilterJenisPeriksa.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   9405
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   9405
      Begin VB.Frame Frame4 
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
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   8895
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Group By"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3735
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Total"
               Height          =   210
               Index           =   5
               Left            =   2640
               TabIndex        =   14
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Tahun"
               Height          =   210
               Index           =   2
               Left            =   1680
               TabIndex        =   13
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Hari"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   0
               Top             =   230
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Bulan"
               Height          =   210
               Index           =   1
               Left            =   840
               TabIndex        =   1
               Top             =   230
               Width           =   735
            End
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   360
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
            OLEDropMode     =   1
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   61145091
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   6600
            TabIndex        =   3
            Top             =   360
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   61145091
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6240
            TabIndex        =   9
            Top             =   480
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraButton 
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
      Top             =   2160
      Width           =   9405
      Begin VB.OptionButton optGroupBy 
         Caption         =   "Jenis Pasien"
         Height          =   210
         Index           =   4
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optGroupBy 
         Caption         =   "Instalasi Asal"
         Height          =   210
         Index           =   3
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Spreadsheet"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   15
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
      Left            =   7560
      Picture         =   "frmFilterJenisPeriksa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmFilterJenisPeriksa.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmFilterJenisPeriksa.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmFilterJenisPeriksa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub cmdCetak_Click()
'If optGroupBy(4).Value = False And optGroupBy(3).Value = False Then MsgBox "Pilih Group By", vbExclamation, "Validasi": Exit Sub
'
'Dim dtbulan As Integer
'Dim dttahun As Integer
''-----------------------------------------------------------------------
'    If optGroupBy(0).Value = True And optGroupBy(3).Value = True Then
'        mdTglAwal = dtpAwal.Value
'        mdTglAkhir = dtpAkhir.Value
'        Select Case strCetak
'             Case "LapKunjunganJenisPeriksa"
'                  strCetak2 = "LapKunjunganJenisPeriksahari"
'                  strCetak3 = "JenisPeriksaBInstalasiAsal"
'                  strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
'                          "WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
'
'             Case "lapkunjunganjenistindakan"
'                 strCetak2 = "LapKunjunganJenisTindakanHari"
'                 strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsal"
'                 strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
'                          "WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
'
'
'
'        End Select
''-------------------------------------------------------------------------
'    ElseIf optGroupBy(0).Value = True And optGroupBy(4).Value = True Then
'        mdTglAwal = dtpAwal.Value
'        mdTglAkhir = dtpAkhir.Value
'            Select Case strCetak
'            Case "LapKunjunganJenisPeriksa"
'                  strCetak2 = "LapKunjunganJenisPeriksahari"
'                  strCetak3 = "JenisPeriksaBJenispasien"
'                  strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
'                          "WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "' and  TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
'
'            Case "lapkunjunganjenistindakan"
'                  strCetak2 = "LapKunjunganJenisTindakanHari"
'                  strCetak3 = "LapKunjunganJenisTindakanBJenisPasienHari"
'                  strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
'                          "WHERE (KdRuanganPelayanan = '" & mstrKdRuangan & "' And TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPelayanan asc"
'
'
'            End Select
''-------------------------------------------------------------------------
'    ElseIf optGroupBy(1).Value = True And optGroupBy(4).Value = True Then
'        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm-01 00:00:00"))
'        dtbulan = CStr(Format(dtpAkhir.Value, "mm"))
'        dttahun = CStr(Format(dtpAkhir.Value, "yyyy"))
'        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm" & "-" & funcHitungHari(dtbulan, dttahun) & " " & "23:59:59"))
'        Select Case strCetak
'        Case "LapKunjunganJenisPeriksa"
'              strCetak2 = "LapKunjunganJenisPeriksabulan"
'              strCetak3 = "JenisPeriksaBJenispasien"
'              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan) AS TglPelayanan, RuanganPelayanan, JenisPeriksa, JenisPasien, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPeriksa " _
'                          & "WHERE (TglPelayanan BETWEEN '" _
'                          & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
'                          & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
'                          " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'        Case "lapkunjunganjenistindakan"
'              strCetak2 = "LapKunjunganJenisTindakanBulan"
'              strCetak3 = "LapKunjunganJenisTindakanBJenisPasienBulan"
'              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan) AS TglPelayanan, RuanganPelayanan, JenisPelayanan, JenisPasien, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPelayanan " _
'                        & "WHERE (TglPelayanan BETWEEN '" _
'                        & Format(mdTglAwal, "yyyy/MM/01 00:00:00") & "' AND '" _
'                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
'                        " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'        End Select
''-------------------------------------------------------------------------
'    ElseIf optGroupBy(1).Value = True And optGroupBy(3).Value = True Then
'        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm-01 00:00:00"))
'        dtbulan = CStr(Format(dtpAkhir.Value, "mm"))
'        dttahun = CStr(Format(dtpAkhir.Value, "yyyy"))
'        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm" & "-" & funcHitungHari(dtbulan, dttahun) & " " & "23:59:59"))
'        Select Case strCetak
'         Case "LapKunjunganJenisPeriksa"
'              strCetak2 = "LapKunjunganJenisPeriksabulan"
'              strCetak3 = "JenisPeriksaBInstalasiAsal"
'              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan) AS TglPelayanan, RuanganPelayanan, JenisPeriksa, InstalasiAsal, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPeriksa " _
'                        & "WHERE (TglPelayanan BETWEEN '" _
'                        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
'                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
'                        " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'
'         Case "lapkunjunganjenistindakan"
'              strCetak2 = "LapKunjunganJenisTindakanBulan"
'              strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsalBulan"
'              strSQL = "SELECT dbo.FB_TakeBlnThn(TglPelayanan)  AS TglPelayanan, RuanganPelayanan, JenisPelayanan, InstalasiAsal, JK, JmlPelayanan  FROM   V_DataKunjunganPasienBJenisPelayanan " _
'                        & "WHERE (TglPelayanan BETWEEN '" _
'                        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
'                        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
'                        " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'         End Select
'
''-------------------------------------------------------------------------
'
'    ElseIf optGroupBy(2).Value = True And optGroupBy(3).Value = True Then
'        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy HH:MM:SS")) 'TglAwal
'        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy HH:MM:SS")) 'TglAkhir
'
'         Select Case strCetak
'         Case "LapKunjunganJenisPeriksa"
'              strCetak2 = "LapKunjunganJenisPeriksaTahun"
'              strCetak3 = "JenisPeriksaBInstalasiAsal"
'              strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
'                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
'                          " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'
'         Case "lapkunjunganjenistindakan"
'              strCetak2 = "LapKunjunganJenisTindakantahun"
'              strCetak3 = "LapKunjunganJenisTindakanBinstalasiAsaltahun"
'              strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
'                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
'                          " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'        End Select
''-------------------------------------------------------------------------
'        ElseIf optGroupBy(2).Value = True And optGroupBy(4).Value = True Then
'        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy HH:MM:SS")) 'TglAwal
'        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy HH:MM:SS")) 'TglAkhir
'
'         Select Case strCetak
'         Case "LapKunjunganJenisPeriksa"
'              strCetak2 = "LapKunjunganJenisPeriksaTahun"
'              strCetak3 = "JenisPeriksaBJenispasien"
'              strSQL = "Select * from V_DataKunjunganPasienBJenisPeriksa " & _
'                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
'                          " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'
'         Case "lapkunjunganjenistindakan"
'              strCetak2 = "LapKunjunganJenisTindakantahun"
'              strCetak3 = "LapKunjunganJenisTindakanBJenisPtahun"
'              strSQL = "Select * from V_DataKunjunganPasienBJenisPelayanan " & _
'                          "WHERE (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
'                          " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"
'        End Select
''--------------------------------------------------------------------------
'
'
'        End If
'    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
'    Call msubRecFO(rs, strSQL)
'    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
'    frmcetakviewer.Show
'End Sub
'
'Private Sub cmdTutup_Click()
'    Unload Me
'End Sub
'
'Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then optGroupBy(0).SetFocus
'End Sub
'
'Private Sub dtpAkhir_Change()
'    dtpAkhir.MaxDate = Now
'End Sub
'
'Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then cmdCetak.SetFocus
'End Sub
'
'Private Sub dtpAwal_Change()
'    dtpAwal.MaxDate = Now
'End Sub
'
'Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dtpAkhir.SetFocus
'End Sub
'
'Private Sub Form_Load()
'    Call centerForm(Me, MDIUtama)
'    Call PlayFlashMovie(Me)
'    With Me
'        .dtpAwal.Value = Now
'        .dtpAkhir.Value = Now
'    End With
'    Call cekOpt
'End Sub
'
'Private Sub cekOpt()
'   If optGroupBy(0).Value = True Then
'      Call optGroupBy_Click(0)
'   ElseIf optGroupBy(1).Value = True Then
'      Call optGroupBy_Click(1)
'   ElseIf optGroupBy(2).Value = True Then
'     Call optGroupBy_Click(2)
'   End If
'End Sub
'
'Private Sub optGroupBy_Click(Index As Integer)
'  Select Case Index
'   Case 0
'        dtpAwal.CustomFormat = "dd MMMM yyyyy"
'        dtpAkhir.CustomFormat = "dd MMMM yyyyy"
'        optGroupBy(1).Value = False
'        optGroupBy(2).Value = False
'        optGroupBy(3).Value = False
'
'   Case 1
'        dtpAkhir.CustomFormat = "MMMM yyyyy"
'        dtpAwal.CustomFormat = "MMMM yyyyy"
'        optGroupBy(0).Value = False
'        optGroupBy(2).Value = False
'        optGroupBy(3).Value = False
'
'   Case 2
'        dtpAkhir.CustomFormat = "yyyyy"
'        dtpAwal.CustomFormat = "yyyyy"
'        optGroupBy(0).Value = False
'        optGroupBy(1).Value = False
'        optGroupBy(3).Value = False
'
'    Case 3
'        dtpAwal.CustomFormat = "dd MMMM yyyyy"
'        dtpAkhir.CustomFormat = "dd MMMM yyyyy"
'        optGroupBy(0).Value = False
'        optGroupBy(1).Value = False
'        optGroupBy(2).Value = False
'   End Select
'End Sub
'
'Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then dtpAwal.SetFocus
'End Sub
'
'
