VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapRKP_KPSK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kunjungan Pasien "
   ClientHeight    =   2910
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
   Icon            =   "frmLapRKP_KPSK.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
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
         Left            =   360
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
               Index           =   3
               Left            =   2640
               TabIndex        =   12
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton optGroupBy 
               Caption         =   "Tahun"
               Height          =   210
               Index           =   2
               Left            =   1680
               TabIndex        =   11
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
            Format          =   437452803
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
            Format          =   437452803
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
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   9405
      Begin VB.CommandButton CmdGrafik 
         Caption         =   "&Grafik"
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
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
      Left            =   7560
      Picture         =   "frmLapRKP_KPSK.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapRKP_KPSK.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapRKP_KPSK.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmLapRKP_KPSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub filter_kriteria()
    'vlidasi tanggal
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    Dim mdtBulan As Integer
    Dim MdtTahun As Integer
    '=====================================================================================================================================================================================================================================================================================================
    'Kriteria Hari

    If optGroupBy(0).Value = True Then

        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusHari"
                strSQL = "Select * from V_DatakunjunganPasienMasukBjenisBstausPasien " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganTriaseStatus"
                strCetak2 = "LapKunjunganTriaseStatusHari"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DatakunjunganPasienMasukBTriaseBstausPasien " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahHari"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBWilayah " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnHari"
                strSQL = "Select * from V_DataKunjunganPasienMasukBstatusBkasusPenyakit " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatushari"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBsetatusBKelas " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusHari"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBRujukan " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusHari"
                strSQL = "Select * from V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and ( TglKeluar BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglKeluar asc"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusHari"
                strSQL = "Select * from V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBjenisTindakan"
                strCetak2 = "LapKunjunganBjenisTindakanHari"
                strSQL = "Select TglPelayanan,RuanganPelayanan,JenisPelayanan,InstalasiAsal,Jk,JmlPelayanan from V_DataKunjunganPasienBJenisPelayanan " & _
                "WHERE (and  kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaHari"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahHari"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"
            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterHari"
                strSQL = "Select * from V_AmbilDataPasienIGDPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"
        End Select
        '===========================================================================================================================================================================================================================================
        'Kriteria Bulan

    ElseIf optGroupBy(1).Value = True Then
        mdTglAwal = CDate(Format(dtpAwal.Value, "yyyy-mm ") & "-01 00:00:00") 'TglAwal

        mdtBulan = CStr(Format(dtpAkhir.Value, "mm"))
        MdtTahun = CStr(Format(dtpAkhir.Value, "yyyy"))
        mdTglAkhir = CDate(Format(dtpAkhir.Value, "yyyy-mm") & "-" & funcHitungHari(mdtBulan, MdtTahun) & " 23:59:59")

        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM   V_DatakunjunganPasienMasukBjenisBstausPasien " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganTriaseStatus"
                strCetak2 = "LapKunjunganTriaseStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DatakunjunganPasienMasukBTriaseBstausPasien " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBWilayah " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBstatusBkasusPenyakit " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran)  AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBsetatusBKelas " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukBsetatusBRujukan " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglKeluar) AS TglKeluar, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " _
                & "WHERE (TglKeluar BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "And kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(TglPendaftaran) AS TglPendaftaran, RuanganPelayanan, Judul, Detail, JK, JmlPasien, KdInstalasi  FROM  V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " _
                & "WHERE (TglPendaftaran BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(tglperiksa) AS tglperiksa, RuanganPelayanan, KdDiagnosa,Diagnosa, StatusKasus, JenisKelamin, JmlKunjungan  FROM  V_DataDiagnosaPasienNew " _
                & "WHERE (tglperiksa BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahBulan"
                strSQL = "SELECT dbo.FB_TakeBlnThn(tglperiksa) AS tglperiksa, RuanganPelayanan, KdDiagnosa, NamaKecamatan, StatusKasus, JenisKelamin, JmlKunjungan  FROM  V_DataDiagnosaPasienNew " _
                & "WHERE (tglperiksa BETWEEN '" _
                & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
                & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "')" & _
                "and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterBulan"
                strSQL = "Select dbo.FB_TakeBlnThn(TglPendaftaran) as tglPendaftaran,Dokter,Judul,Detail,JK,JmlPasien,KdRuangan from V_AmbilDataPasienIGDPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"
        End Select
        '==================================================================================================================================================================================================================================================================================
        'Kriteria Tahun

    ElseIf optGroupBy(2).Value = True Then
        mdTglAwal = CDate("01-01-" & Format(dtpAwal.Value, "yyyy 00:00:00")) 'TglAwal
        mdTglAkhir = CDate("31-12-" & Format(dtpAkhir.Value, "yyyy 23:59:59")) 'TglAkhir

        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusTahun"
                strSQL = "Select * from V_DatakunjunganPasienMasukBjenisBstausPasien " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganTriaseStatus"
                strCetak2 = "LapKunjunganTriaseStatusTahun"
                strSQL = "Select * from V_DatakunjunganPasienMasukBTriaseBstausPasien " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBWilayah " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBstatusBkasusPenyakit " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBKelas " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBRujukan " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and ( TglKeluar BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "  order by tglKeluar asc"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusTahun"
                strSQL = "Select * from V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " order by TglPendaftaran asc"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaTahun"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (tglperiksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahTahun"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (tglperiksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPerDokter"
                strCetak2 = "LapKunjunganPerDokterTahun"
                strSQL = "Select * from V_AmbilDataPasienIGDPerDokter " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  KdRuangan='" & mstrKdRuangan & "'"

        End Select
        '=========================================================================================================================================================================================================================================================================================================
        'Kriteria Total

    ElseIf optGroupBy(3).Value = True Then
        Select Case strCetak
            Case "LapKunjunganJenisStatus"
                strCetak2 = "LapKunjunganJenisStatusTotal"
                strSQL = "Select * from V_DatakunjunganPasienMasukBjenisBstausPasien " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganTriaseStatus"
                strCetak2 = "LapKunjunganTriaseStatusTotal"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DatakunjunganPasienMasukBTriaseBstausPasien " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBwilayah"
                strCetak2 = "LapKunjunganBwilayahTotal"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBWilayah " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganSt_PnyktPsn"
                strCetak2 = "LapKunjunganSt_PnyktPsnTotal"
                strSQL = "Select * from V_DataKunjunganPasienMasukBstatusBkasusPenyakit " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKelasStatus"
                strCetak2 = "LapKunjunganKelasStatusTotal"
                strSQL = "Select TglPendaftaran,RuanganPelayanan,Judul,Detail,Jk,JmlPasien from V_DataKunjunganPasienMasukBsetatusBKelas " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganRujukanBStatus"
                strCetak2 = "LapKunjunganRujukanBStatusTotal"
                strSQL = "Select * from V_DataKunjunganPasienMasukBsetatusBRujukan " & _
                "WHERE (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                "and  kdRuanganPelayanan='" & mstrKdRuangan & "'"

            Case "LapKunjunganKonPulang_Status"
                strCetak2 = "LapKunjunganKonPulang_StatusTotal"
                strSQL = "Select * from V_DataKunjunganPasienKeluarBKondisiPulang_Bstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and ( TglKeluar BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglKeluar asc"

            Case "LapKunjunganJenisOperasi_Status"
                strCetak2 = "LapKunjunganJenisOperasi_StatusTotal"
                strSQL = "Select * from V_DataKunjunganPasienMasukIBSBJenisOperasiBstatus " & _
                "WHERE (kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPendaftaran BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBjenisTindakan"
                strCetak2 = "LapKunjunganBjenisTindakanTotal"
                strSQL = "Select TglPelayanan,RuanganPelayanan,JenisPelayanan,InstalasiAsal,Jk,JmlPelayanan from V_DataKunjunganPasienBJenisPelayanan " & _
                "WHERE (and  kdRuanganPelayanan='" & mstrKdRuangan & "') and (TglPelayanan BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102)) order by TglPendaftaran asc"

            Case "LapKunjunganBDiagnosa"
                strCetak2 = "LapKunjunganBDiagnosaTotal"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"

            Case "LapKunjunganPasienBDiagnosaWilayah"
                strCetak2 = "LapKunjunganPasienBDiagnosaWilayahTotal"
                strSQL = "Select * from V_DataDiagnosaPasienNew " & _
                "WHERE (TglPeriksa BETWEEN convert(datetime,'" & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' ,102) AND convert (datetime,'" & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "',102))" & _
                " and  kdRuangan='" & mstrKdRuangan & "'"
        End Select

    End If

End Sub

Private Sub cmdCetak_Click()
    Dim pesan As VbMsgBoxResult
    On Error GoTo hell
    Call filter_kriteria
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    FrmCetakLapKunjunganPasien.Show
    Exit Sub
hell:

End Sub

Private Sub cmdgrafik_Click()
    Call filter_kriteria
    If ValidasiTanggal(dtpAwal, dtpAkhir) = False Then Exit Sub
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then MsgBox "Data Tidak Ada", vbExclamation, "Validasi": Exit Sub
    FrmCetakLaporandalamBentukGrafik.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcInstalasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then optGroupBy(0).SetFocus
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCetak.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
    End With

    Call cekOpt
End Sub

Private Sub cekOpt()
    If optGroupBy(0).Value = True Then
        Call optGroupBy_Click(0)
    ElseIf optGroupBy(1).Value = True Then
        Call optGroupBy_Click(1)
    ElseIf optGroupBy(2).Value = True Then
        Call optGroupBy_Click(2)
    End If
End Sub

Private Sub optGroupBy_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(1).Value = False
            optGroupBy(2).Value = False
            optGroupBy(3).Value = False

        Case 1
            dtpAkhir.CustomFormat = "MMMM yyyyy"
            dtpAwal.CustomFormat = "MMMM yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(2).Value = False
            optGroupBy(3).Value = False

        Case 2
            dtpAkhir.CustomFormat = "yyyyy"
            dtpAwal.CustomFormat = "yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(1).Value = False
            optGroupBy(3).Value = False
        Case 3
            dtpAwal.CustomFormat = "dd MMMM yyyyy"
            dtpAkhir.CustomFormat = "dd MMMM yyyyy"
            optGroupBy(0).Value = False
            optGroupBy(1).Value = False
            optGroupBy(2).Value = False
    End Select
End Sub

Private Sub optGroupBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAwal.SetFocus
End Sub

