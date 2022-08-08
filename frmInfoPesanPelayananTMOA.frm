VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInfoPesanPelayananTMOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Pesan Pelayanan Ruangan"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfoPesanPelayananTMOA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   13830
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   13815
      Begin VB.Frame Frame5 
         Caption         =   "Pelayanan"
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optTM 
            Caption         =   "Tindakan Medis"
            Height          =   375
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optOA 
            Caption         =   "Obat Resep"
            Height          =   375
            Left            =   2040
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status Verifikasi"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton optSudah 
            Caption         =   "Sudah"
            Height          =   375
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optBelum 
            Caption         =   "Belum"
            Height          =   375
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tgl. Pemesanan"
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
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Width           =   6975
         Begin VB.CommandButton cmdTampilkan 
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
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpTglAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   128909315
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpTglAkhir 
            Height          =   375
            Left            =   4200
            TabIndex        =   7
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy HH:mm"
            Format          =   128909315
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3720
            TabIndex        =   15
            Top             =   315
            Width           =   255
         End
      End
      Begin MSDataGridLib.DataGrid dgInfoPesanPelayanan 
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   9763
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   7800
      Width           =   13815
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   520
         Left            =   10560
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ceta&k Struk"
         Height          =   520
         Left            =   8880
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   520
         Left            =   12240
         TabIndex        =   11
         Top             =   240
         Width           =   1455
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
      Left            =   12000
      Picture         =   "frmInfoPesanPelayananTMOA.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInfoPesanPelayananTMOA.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInfoPesanPelayananTMOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCetak_Click()
If dgInfoPesanPelayanan.ApproxCount = 0 Then Exit Sub
mstrNoPen = dgInfoPesanPelayanan.Columns("NoPendaftaran").Value
If optOA.Value = True Then
    If optSudah.Value = True Then Exit Sub
    
    mdTglAwal = dtpTglAwal.Value
    mdTglAkhir = dtpTglAkhir.Value
    mstrNoPen = dgInfoPesanPelayanan.Columns("NoPendaftaran").Value
    mstrKdRuanganORS = dgInfoPesanPelayanan.Columns("KdRuanganTujuan").Value
    strNamaRuangan = dgInfoPesanPelayanan.Columns("RuanganTujuan").Value
    mstrNama = dgInfoPesanPelayanan.Columns("DokterOrder")
    
    mstrNoOrder = dgInfoPesanPelayanan.Columns("NoOrder").Value
    
    strCetak2 = "OA"
    
    frm_cetak_RincianBiayaKonsul.Show

End If
If optTM.Value = True Then
    mdTglAwal = dtpTglAwal.Value
    mdTglAkhir = dtpTglAkhir.Value
    mstrKdRuanganORS = dgInfoPesanPelayanan.Columns("KdRuanganTujuan").Value

    strNStsCITO = dgInfoPesanPelayanan.Columns("StatusCito")

    strNamaRuangan = dgInfoPesanPelayanan.Columns("RuanganTujuan").Value
    
    If dgInfoPesanPelayanan.Columns("DokterOrder") = "" Then
        mstrNama = ""
    Else
        mstrNama = dgInfoPesanPelayanan.Columns("DokterOrder").Value
    End If
    mstrKdKelas = dgInfoPesanPelayanan.Columns("KdKelasAkhir").Value
    mstrNoOrder = dgInfoPesanPelayanan.Columns("NoOrder").Value
    
    If optSudah.Value = True Then
        strCetak = "1"
    ElseIf optBelum.Value = True Then
        strCetak = "0"
    End If
    strCetak2 = "TM"
    frm_cetak_RincianBiayaKonsul.Show

End If
End Sub

Private Sub cmdHapus_Click()
On Error GoTo hell
If dgInfoPesanPelayanan.ApproxCount = 0 Then Exit Sub
If optSudah.Value = True Then Exit Sub
If MsgBox("Anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
If optTM.Value = True Then
    strSQL = "Delete DetailOrderPelayananTM where NoOrder='" & dgInfoPesanPelayanan.Columns("NoOrder").Value & "' AND NoPendaftaran='" & dgInfoPesanPelayanan.Columns("NoPendaftaran").Value & "' AND KdPelayananRS='" & dgInfoPesanPelayanan.Columns("KdPelayananRS").Value & "' "
    Call msubRecFO(rs, strSQL)
'    strSQL = "Delete StrukOrder where NoOrder='" & dgInfoPesanPelayanan.Columns("NoOrder").Value & "'"
'    Call msubRecFO(rs, strSQL)
'
    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
    
    Call Add_HistoryLoginActivity("Delete_OrderPelayananTM")
    subLoadData
End If
If optOA.Value = True Then
    strSQL = "Delete DetailOrderPelayananOARacikanTemp where NoRacikan='" & dgInfoPesanPelayanan.Columns("NoRacikan").Value & "' AND KdBarang='" & dgInfoPesanPelayanan.Columns("KdBarang").Value & "' "
    Call msubRecFO(rs, strSQL)
    strSQL = "Delete DetailOrderPelayananOARacikan where NoRacikan='" & dgInfoPesanPelayanan.Columns("NoRacikan").Value & "' AND KdBarang='" & dgInfoPesanPelayanan.Columns("KdBarang").Value & "' "
    Call msubRecFO(rs, strSQL)
    strSQL = "Delete DetailOrderPelayananOA where NoOrder='" & dgInfoPesanPelayanan.Columns("NoOrder").Value & "' AND NoPendaftaran='" & dgInfoPesanPelayanan.Columns("NoPendaftaran").Value & "' AND KdBarang='" & dgInfoPesanPelayanan.Columns("KdBarang").Value & "' "
    Call msubRecFO(rs, strSQL)
    
'    strSQL = "Delete StrukOrder where NoOrder='" & dgInfoPesanPelayanan.Columns("NoOrder").Value & "'"
'    Call msubRecFO(rs, strSQL)
    
    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
    
    Call Add_HistoryLoginActivity("Delete_OrderPelayananOA")
    subLoadData
End If
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTampilkan_Click()
    Call subLoadData
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgInfoPesanPelayanan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgInfoPesanPelayanan
    WheelHook.WheelHook dgInfoPesanPelayanan
End Sub

Private Sub dtpTglAkhir_Change()
    dtpTglAkhir.MaxDate = Now
End Sub

Private Sub dtpTglAwal_Change()
    dtpTglAwal.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    optBelum.Value = True
    optTM.Value = True
    dtpTglAkhir.Value = Format(Now, "dd MMMM yyyy 23:59:59")
    dtpTglAwal.Value = Format(Now, "dd MMMM yyyy 00:00:00")
    Call subLoadData

End Sub

Sub subLoadData()
    On Error GoTo hell
    Dim i As Integer

    If optTM.Value = True Then
        If optBelum.Value = True Then
            strSQL = "Select NoOrder,NamaPelayanan,JmlPelayanan,StatusCito,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdPelayananRS,KdRuanganTujuan,KdKelasAkhir from V_DaftarDetailOrderTM where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is null"
        ElseIf optSudah.Value = True Then
            strSQL = "Select NoOrder,NamaPelayanan,JmlPelayanan,StatusCito,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdPelayananRS,KdRuanganTujuan,KdKelasAkhir from V_DaftarDetailOrderTM where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is not null"
        End If
    ElseIf optOA.Value = True Then
        If optBelum.Value = True Then
            If bolStatusFIFO = True Then
            strSQL = "Select  distinct NoOrder,JenisBarang,NamaBarang,JmlBarang,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdBarang,NoRacikan,KdRuanganTujuan,NoRacikan from V_DaftarDetailOrderOAFIFO where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is null"
            Else
            strSQL = "Select  distinct NoOrder,JenisBarang,NamaBarang,JmlBarang,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdBarang,NoRacikan,KdRuanganTujuan,NoRacikan from V_DaftarDetailOrderOA where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is null"
            End If
        ElseIf optSudah.Value = True Then
            If bolStatusFIFO = True Then
            strSQL = "Select distinct NoOrder,JenisBarang,NamaBarang,JmlBarang,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdBarang,NoRacikan,KdRuanganTujuan,NoRacikan from V_DaftarDetailOrderOAFIFO where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is not null"
            Else
            strSQL = "Select distinct NoOrder,JenisBarang,NamaBarang,JmlBarang,Satuan,NoPendaftaran,NoCM,NamaPasien,RuanganTujuan,DokterOrder,UserOrder,KdBarang,NoRacikan,KdRuanganTujuan,NoRacikan from V_DaftarDetailOrderOA where TglOrder Between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "' and KdRuangan='" & mstrKdRuangan & "' and NoRiwayat is not null"
            End If
        End If
    End If

    Set rs = Nothing
    Call msubRecFO(rs, strSQL)

    Set dgInfoPesanPelayanan.DataSource = rs

    With dgInfoPesanPelayanan

        If optTM.Value = True Then
            .Columns("NamaPelayanan").Width = 2000
            .Columns("JmlPelayanan").Width = 800
            .Columns("JmlPelayanan").Caption = "Jumlah"

            .Columns("StatusCito").Width = 800
            .Columns("NoPendaftaran").Width = 1350
            .Columns("NoCM").Width = 1500
'            .Columns("NoCM").Caption = "No. Rekam Medis"
            .Columns("NamaPasien").Width = 2500
            .Columns("RuanganTujuan").Width = 2200
            .Columns("DokterOrder").Width = 2000
            .Columns("UserOrder").Width = 2000
            .Columns("NoOrder").Width = 1500
            .Columns("KdPelayananRS").Width = 0
            .Columns("KdRuanganTujuan").Width = 0
            .Columns("KdKelasAkhir").Width = 0
        ElseIf optOA.Value = True Then
            .Columns("JenisBarang").Width = 1500
            .Columns("NamaBarang").Width = 2000
            '.Columns("NamaAsal").Width = 1500
            .Columns("JmlBarang").Width = 800
            .Columns("JmlBarang").Caption = "Jumlah"
'            .Columns("HargaFIFO").Width = 1500
'            .Columns("HargaFIFO").Alignment = dbgRight
            .Columns("Satuan").Width = 1000
            .Columns("NoPendaftaran").Width = 1350
            .Columns("NoCM").Width = 1000
'            .Columns("NoCM").Caption = "No. Rekam Medis"
            .Columns("NamaPasien").Width = 2500
            .Columns("RuanganTujuan").Width = 2200
            .Columns("DokterOrder").Width = 2000
            .Columns("UserOrder").Width = 2000
            .Columns("NoOrder").Width = 1500
            .Columns("KdBarang").Width = 0
            .Columns("NoRacikan").Width = 0
            .Columns("KdRuanganTujuan").Width = 0
        End If
    End With

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub optBelum_Click()
cmdHapus.Enabled = True
cmdCetak.Enabled = True
subLoadData
End Sub

Private Sub optOA_Click()
If optSudah.Value = True Then cmdCetak.Enabled = False
subLoadData
End Sub

Private Sub optSudah_Click()
cmdHapus.Enabled = False
If optOA.Value = True Then cmdCetak.Enabled = False
subLoadData
End Sub

Private Sub optTM_Click()
cmdCetak.Enabled = True
subLoadData
End Sub
