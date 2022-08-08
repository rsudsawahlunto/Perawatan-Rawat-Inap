VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmSensusPelayanan_HAJI_DataGrid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Sensus Pelayanan"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSensusPelayanan_HAJI_DataGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   15045
   Begin VB.Frame Frame4 
      Caption         =   "Rincian Total Komponen Tarif"
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
      TabIndex        =   17
      Top             =   5760
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtGTBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtGTKomponen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtGTJasaKomponen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtTotalBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtTotalKomponen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtJasaKomponen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label lblTotalBiaya 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Total Biaya"
         Height          =   210
         Left            =   5280
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "Total Komponen"
         Height          =   210
         Left            =   3600
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblKonponenTarif 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Komponen Tarif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1965
         TabIndex        =   23
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblKelompokPasien 
         AutoSize        =   -1  'True
         Caption         =   "Kelompok Pasien"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblJenisPasien 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   660
         Width           =   525
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
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
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   7320
      Width           =   15015
      Begin VB.CheckBox Check1 
         Caption         =   "Total Biaya"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   495
         Left            =   11280
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13200
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   15015
      Begin VB.Frame Frame3 
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
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   5775
         Begin VB.CommandButton cmdTampilkanTemp 
            Caption         =   "&Cari"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   3
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   54722563
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   4
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
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   54722563
            UpDown          =   -1  'True
            CurrentDate     =   38373
         End
         Begin VB.Label Label1 
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
            Left            =   3120
            TabIndex        =   13
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataListLib.DataCombo dcStatusBayar 
         Height          =   360
         Left            =   4560
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo dcJenisLaporan 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo dcKomponenTarif 
         Height          =   360
         Left            =   6840
         TabIndex        =   2
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Tarif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   6840
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Status Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Laporan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   30
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmSensusPelayanan_HAJI_DataGrid.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   13200
      Picture         =   "frmSensusPelayanan_HAJI_DataGrid.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmSensusPelayanan_HAJI_DataGrid.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmSensusPelayanan_HAJI_DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long

Private Sub Check1_Click()
Dim curGTJasaKomponen As Currency
Dim curGTKomponen As Currency
Dim curGTTotalBiaya As Currency

    lblJenisPasien(0).Caption = ""
    txtJasaKomponen(0).Text = 0
    txtTotalKomponen(0).Text = 0
    txtTotalBiaya(0).Text = 0
    
    If Check1.Value = vbUnchecked Then
        curGTJasaKomponen = 0: curGTKomponen = 0: curGTTotalBiaya = 0
        For i = 1 To lblJenisPasien.Count - 1
            Frame4.Top = Frame4.Top + txtJasaKomponen(i - 1).Height
            Frame4.Height = Frame4.Height - txtJasaKomponen(i - 1).Height
            
            Unload lblJenisPasien(i)
            Unload txtJasaKomponen(i)
            Unload txtTotalKomponen(i)
            Unload txtTotalBiaya(i)
        Next i
        txtGTJasaKomponen.Top = txtGTJasaKomponen.Top + txtJasaKomponen(0).Height + 120
        txtGTKomponen.Top = txtGTJasaKomponen.Top
        txtGTBiaya.Top = txtGTJasaKomponen.Top
        lblGrandTotal.Top = txtGTJasaKomponen.Top + 60
        
        Frame4.Visible = False
        Exit Sub
    End If
    
    lblKonponenTarif.Caption = dcKomponenTarif.Text
    strSQL = " SELECT JenisPasien, SUM(Harga) AS Harga, SUM(Total) AS Total, SUM(TotalBiaya) AS TotalBiaya " & _
        " FROM V_SensusPendapatan " & _
        " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
        " GROUP BY JenisPasien"
    Call msubRecFO(rs, strSQL)
      
    For i = 1 To rs.RecordCount - 1
        Load txtJasaKomponen(i)
        txtJasaKomponen(i).Top = txtJasaKomponen(i - 1).Top + txtJasaKomponen(i - 1).Height
        txtJasaKomponen(i).Visible = True
    
        Load lblJenisPasien(i)
        lblJenisPasien(i).Top = lblJenisPasien(i - 1).Top + txtJasaKomponen(i - 1).Height
        lblJenisPasien(i).Visible = True
    
        Load txtTotalKomponen(i)
        txtTotalKomponen(i).Top = txtTotalKomponen(i - 1).Top + txtTotalKomponen(i - 1).Height
        txtTotalKomponen(i).Visible = True
    
        Load txtTotalBiaya(i)
        txtTotalBiaya(i).Top = txtTotalBiaya(i - 1).Top + txtTotalBiaya(i - 1).Height
        txtTotalBiaya(i).Visible = True
        
        Frame4.Top = Frame4.Top - txtJasaKomponen(i - 1).Height
        Frame4.Height = Frame4.Height + txtJasaKomponen(i - 1).Height
    Next i
    
    
    For i = 0 To rs.RecordCount - 1
        lblJenisPasien(i).Caption = rs(0).Value
        txtJasaKomponen(i).Text = Format(rs(1).Value, "#,###"): curGTJasaKomponen = curGTJasaKomponen + rs(1).Value
        txtTotalKomponen(i).Text = Format(rs(2).Value, "#,###"): curGTKomponen = curGTKomponen + rs(2).Value
        txtTotalBiaya(i).Text = Format(rs(3).Value, "#,###"): curGTTotalBiaya = curGTTotalBiaya + rs(3).Value
        rs.MoveNext
    Next i
    
    txtGTJasaKomponen.Top = txtJasaKomponen(txtJasaKomponen.UBound).Top + txtJasaKomponen(txtJasaKomponen.UBound).Height + 120
    txtGTKomponen.Top = txtGTJasaKomponen.Top
    txtGTBiaya.Top = txtGTJasaKomponen.Top
    lblGrandTotal.Top = txtGTJasaKomponen.Top + 60
    
    txtGTJasaKomponen.Text = IIf(curGTJasaKomponen = 0, 0, Format(curGTJasaKomponen, "#,###"))
    txtGTKomponen.Text = IIf(curGTJasaKomponen = 0, 0, Format(curGTJasaKomponen, "#,###"))
    txtGTBiaya.Text = IIf(curGTTotalBiaya = 0, 0, Format(curGTTotalBiaya, "#,###"))
    
    Frame4.Visible = True
End Sub

Private Sub cmdCetak_Click()
On Error GoTo errload

'    If Periksa("datacombo", dcJenisLaporan, "Jenis laporan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcStatusBayar, "Status bayar kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKomponenTarif, "Komponen tarif kosong") = False Then Exit Sub
    
    strSQL = " SELECT * FROM V_D_LaporanSensusPelayanan " & _
             " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "'" & _
             " AND KdRuangan = '" & mstrKdRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbExclamation, "Validasi"
        Exit Sub
    End If

    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
'    mstrKdJenisLaporan = dcJenisLaporan.BoundText
    mstrNamaKomponenTarif = dcKomponenTarif.Text
    mstrStatusBayar = dcStatusBayar.Text

    Set frmCetakSensusPelayanan = Nothing
    frmCetakSensusPelayanan.Show
    cmdCetak.Enabled = True

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdTampilkanTemp_Click()
    On Error GoTo errTampilkan
'    If Periksa("datacombo", dcJenisLaporan, "Jenis laporan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcStatusBayar, "Status bayar kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKomponenTarif, "Komponen tarif kosong") = False Then Exit Sub
    
    Select Case dcStatusBayar.BoundText
        Case "01" 'Belum
            mstrFilterData = " AND (NoStruk IS NULL)"
        Case "02" 'sudah
            mstrFilterData = " AND (NoStruk IS NOT NULL)"
        Case Else 'total
            mstrFilterData = ""
    End Select
   
    Call subLoadDataGrid
   
    Check1.Value = vbUnchecked
Exit Sub
errTampilkan:
    msubPesanError
End Sub

'Private Sub dcJenisLaporan_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dcStatusBayar.SetFocus
'End Sub

Private Sub dcKomponenTarif_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then dtpAwal.SetFocus
End Sub

Private Sub dcStatusBayar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then dcKomponenTarif.SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdTampilkanTemp.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errFormLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    Me.Caption = "Medifirst2000 - Laporan Sensus Pelayanan"
    dtpAwal.Value = Now
    dtpAkhir.Value = Now
    
    Call subLoadDC
    Call cmdTampilkanTemp_Click
Exit Sub
errFormLoad:
    msubPesanError
End Sub

Private Sub subLoadDC()
On Error GoTo errload
'    strSQL = "SELECT * FROM JenisLaporanKasir"
'    Call msubDcSource(dcJenisLaporan, rs, strSQL)
'    If Not rs.EOF Then dcJenisLaporan.Text = rs(1).Value
    
    strSQL = "SELECT * FROM StatusBayar"
    Call msubDcSource(dcStatusBayar, rs, strSQL)
    If Not rs.EOF Then dcStatusBayar.Text = rs(1).Value
    
    'View U/ RI
    strSQL = "SELECT * FROM V_KomponenTarifJPOM"
    Call msubDcSource(dcKomponenTarif, rs, strSQL)
    If Not rs.EOF Then dcKomponenTarif.Text = rs(1).Value
    
    Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub subLoadDataGrid()
    strSQL = " SELECT JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya " & _
        " FROM V_SensusPendapatan " & _
        " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
        " GROUP BY JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya "
    msubRecFO rs, strSQL
    Set DataGrid1.DataSource = rs
    DataGrid1.Columns("Harga").Caption = dcKomponenTarif.Text
    DataGrid1.Columns(dcKomponenTarif.Text).Alignment = dbgRight
    DataGrid1.Columns(dcKomponenTarif.Text).NumberFormat = "#,###"
    
    DataGrid1.Columns("Jumlah").Alignment = dbgCenter
    
    DataGrid1.Columns("Total").Alignment = dbgRight
    DataGrid1.Columns("Total").NumberFormat = "#,###"
    
    DataGrid1.Columns("TotalBiaya").Alignment = dbgRight
    DataGrid1.Columns("TotalBiaya").NumberFormat = "#,###"
    
    With DataGrid1
    .Columns(0).Width = 1500 'Jenis Pasien
    .Columns(1).Width = 1150 'No Register
    .Columns(2).Width = 800 'No. CM
    .Columns(3).Width = 2000 'Nama Pasien
    .Columns(4).Width = 1700 'Tgl. Pelayanan
    .Columns(5).Width = 2050 'Jenis Pelayanan
    .Columns(6).Width = 2050 'Nama Pelayanan
    .Columns(7).Width = 1200 'Kelas
    .Columns(8).Width = 1700 'Komponen Tarif...
    .Columns(9).Width = 600 'Jml
    .Columns(10).Width = 1450 'Total
    .Columns(11).Width = 1450 'Total Biaya
End With

End Sub

'Private Function subLoadSemuaDataGrid(f_KdJenisLaporan As String) As String
'    Select Case f_KdJenisLaporan
'        Case "01" 'tindakan
'            subLoadSemuaDataGrid = " SELECT JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya " & _
'                " FROM V_SensusPendapatanTindakan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya "
'        Case "02" 'Obat Alkes
'            subLoadSemuaDataGrid = " SELECT JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya " & _
'                " FROM V_SensusPendapatanObatAlkesAll " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya "
'        Case "03" 'Total
'            subLoadSemuaDataGrid = " SELECT JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya " & _
'                " FROM V_SensusPendapatan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya "
'        Case Else
'            subLoadSemuaDataGrid = " SELECT JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya " & _
'                " FROM V_SensusPendapatan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien, NoRegistrasi, NoCM, NamaPasien, TglPelayanan, JenisPelayanan, NamaPelayanan, Kelas, Harga, Jumlah, Total, TotalBiaya "
'    End Select
'End Function

'Private Function subLoadSubTotalGrid(f_KdJenisLaporan As String) As String
'    Select Case f_KdJenisLaporan
'        Case "01" 'tindakan
'            subLoadSubTotalGrid = " SELECT JenisPasien, SUM(Harga) AS Harga, SUM(Total) AS Total, SUM(TotalBiaya) AS TotalBiaya " & _
'                " FROM V_SensusPendapatanTindakan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien"
'        Case "02" 'Obat Alkes
'            subLoadSubTotalGrid = " SELECT JenisPasien, SUM(Harga) AS Harga, SUM(Total) AS Total, SUM(TotalBiaya) AS TotalBiaya " & _
'                " FROM V_SensusPendapatanObatAlkesAll " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien"
'        Case "03" 'Total
'            subLoadSubTotalGrid = " SELECT JenisPasien, SUM(Harga) AS Harga, SUM(Total) AS Total, SUM(TotalBiaya) AS TotalBiaya " & _
'                " FROM V_SensusPendapatan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien"
'        Case Else
'            subLoadSubTotalGrid = " SELECT JenisPasien, SUM(Harga) AS Harga, SUM(Total) AS Total, SUM(TotalBiaya) AS TotalBiaya " & _
'                " FROM V_SensusPendapatan " & _
'                 " WHERE TglPelayanan BETWEEN '" & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "' AND KdRuangan = '" & mstrKdRuangan & "' AND KomponenTarif = '" & dcKomponenTarif.Text & "' " & mstrFilterData & "" & _
'                 " GROUP BY JenisPasien"
'    End Select
'End Function


