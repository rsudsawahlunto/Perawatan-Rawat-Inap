VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Medifirst2000 -  Perawatan_Rawat_Inap (InPatient)"
   ClientHeight    =   8100
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15180
   Icon            =   "MDIFrm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrm1.frx":0CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDPrinter 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "27/10/2014"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:40"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6059
            MinWidth        =   6068
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnberkas 
      Caption         =   "&Berkas"
      Begin VB.Menu mnudata 
         Caption         =   "Data"
         Begin VB.Menu mnucdp 
            Caption         =   "Daftar Pasien Rawat Inap"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuseppm 
            Caption         =   "-"
         End
         Begin VB.Menu mnudpm 
            Caption         =   "Daftar Pasien Meninggal"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnusepdpm 
            Caption         =   "-"
         End
         Begin VB.Menu mnudpl 
            Caption         =   "Daftar Pasien Lama"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu sepdpl 
            Caption         =   "-"
         End
         Begin VB.Menu MDaftarDokumenRekamMedis 
            Caption         =   "Daftar Dokumen Rekam Medis"
            Shortcut        =   {F12}
         End
         Begin VB.Menu lnDok 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBDataDiag 
            Caption         =   "Diagnosa Ruangan"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnusepdr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDaftarBayiLahir 
            Caption         =   "Daftar Bayi Lahir"
            Shortcut        =   +{F9}
         End
         Begin VB.Menu mnuJenisPersalinanEventBayi 
            Caption         =   "Jenis Persalinan dan Event Bayi"
         End
         Begin VB.Menu lnjenisprsalinanNevenBayi 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnupp 
            Caption         =   "Paket Pelayanan"
            Shortcut        =   ^D
            Visible         =   0   'False
         End
         Begin VB.Menu mnuline 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMastDiagnosaKeperawatan 
            Caption         =   "Master Diagnosa Keperawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDetailDiagnosaKeperawatan 
            Caption         =   "Detail Diagnosa Keperawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu MPenyebabDiagosaKeperawatan 
            Caption         =   "Penyebab Diagosa Keperawatan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTujuanNRencanaTindakan 
            Caption         =   "Tujuan && Rencana Tindakan"
            Visible         =   0   'False
         End
         Begin VB.Menu LInformasiTarifPelayanan 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuClosingDataPelayananTMOAApotik 
            Caption         =   "Informasi Data Pelayanan TMOAApotik"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuseppp 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu MInformasiTarifPelayanan 
            Caption         =   "Informasi Tarif Pelayanan"
         End
      End
      Begin VB.Menu mnuseptp 
         Caption         =   "-"
      End
      Begin VB.Menu mSettingPrinter 
         Caption         =   "Setting Printer"
         Shortcut        =   ^P
      End
      Begin VB.Menu mGantiKataKunci 
         Caption         =   "Ganti Kata Kunci"
         Shortcut        =   ^G
      End
      Begin VB.Menu mspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnlogout 
         Caption         =   "Log Off"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnSelesai 
         Caption         =   "Keluar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuInformasi 
      Caption         =   "&Informasi"
      Begin VB.Menu mnMonitoring 
         Caption         =   "Monitoring Pembayaran Pasien"
      End
      Begin VB.Menu mnuPesanPelayananTMOA 
         Caption         =   "Daftar Pesan Pelayanan TMOA"
      End
      Begin VB.Menu mnuDaftarPengirimanDarah 
         Caption         =   "Daftar Pengiriman Darah"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuivt 
      Caption         =   "In&ventory"
      Begin VB.Menu mnupb 
         Caption         =   "Pemesanan Barang"
      End
      Begin VB.Menu mnuPakaiBahan 
         Caption         =   "Pemakaian Bahan && Alat"
         Visible         =   0   'False
      End
      Begin VB.Menu batasinvento 
         Caption         =   "-"
      End
      Begin VB.Menu MBarangMedis 
         Caption         =   "Barang Medis"
         Begin VB.Menu mnusb 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MRekapitulasiTransaksiBarang 
            Caption         =   "Rekapitulasi Transaksi Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu MClosingStok 
            Caption         =   "Closing Stok"
            Begin VB.Menu MCetakFormulirStok 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu MStokOpname 
               Caption         =   "Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaan 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu LClosingStok 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarang 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu MInformasiPemakaianBarang 
            Caption         =   "Informasi Pemakaian Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu MLaporanSaldoBarang 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MBarangNonMedis 
         Caption         =   "Barang Non Medis"
         Begin VB.Menu MStokBarangNonMedis 
            Caption         =   "Stok Barang"
         End
         Begin VB.Menu MKondisiBarangNonMedis 
            Caption         =   "Kondisi Barang"
         End
         Begin VB.Menu MMutasiBarangNonMedis 
            Caption         =   "Mutasi Barang"
         End
         Begin VB.Menu MRekapitulasiTransaksiBarangNonMedis 
            Caption         =   "Rekapitulasi Transaksi Barang"
            Visible         =   0   'False
         End
         Begin VB.Menu MClosingStokNonMedis 
            Caption         =   "Closing Stok"
            Begin VB.Menu MCetakFormulirStokNonMedis 
               Caption         =   "Cetak Lembar Input"
            End
            Begin VB.Menu MStokOpnameNonMedis 
               Caption         =   "Stok Opname"
            End
            Begin VB.Menu MNilaiPersediaanNonMedis 
               Caption         =   "Nilai Persediaan"
            End
         End
         Begin VB.Menu mnusepsb 
            Caption         =   "-"
         End
         Begin VB.Menu MInformasiPemesananPenerimaanBarangNonMedis 
            Caption         =   "Informasi Pemesanan && Penerimaan Barang"
         End
         Begin VB.Menu MLaporanSaldoBarangNonMedis 
            Caption         =   "Laporan Saldo Barang"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuLap 
      Caption         =   "&Laporan"
      Begin VB.Menu mnubrp 
         Caption         =   "Buku Register Pasien"
      End
      Begin VB.Menu mnLPBPP 
         Caption         =   "Laporan Buku Register Pelayanan Pasien"
      End
      Begin VB.Menu LapSensusHarian 
         Caption         =   "Laporan Sensus Harian"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLap1 
         Caption         =   "Laporan Sensus Harian Pasien yang tersisa di ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLap2 
         Caption         =   "Laporan Sensus Harian Pasien yang pulang di ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLap3 
         Caption         =   "Laporan Pendapatan Ruangan"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaa 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu BSJ 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Jenis"
         Visible         =   0   'False
      End
      Begin VB.Menu RKPR 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Rujukan"
         Visible         =   0   'False
      End
      Begin VB.Menu BSDJP 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Kasus Penyakit"
         Visible         =   0   'False
      End
      Begin VB.Menu BSDKP 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Kondisi Pulang"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBKel 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Status dan Kelas"
         Visible         =   0   'False
      End
      Begin VB.Menu ooooooooooo 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBW 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBD 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Diagnosa"
         Visible         =   0   'False
      End
      Begin VB.Menu RPBWD 
         Caption         =   "Rekapitulasi Pasien Berdasarkan Wilayah Diagnosa"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepk1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnur10bp 
         Caption         =   "Rekapitulasi 10 Besar Penyakit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudsmp 
         Caption         =   "Data Surveilens Morbiditas Pasien"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepkp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnurkri 
         Caption         =   "Rekapitulasi Kamar Rawat Inap"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuiprs 
         Caption         =   "Indikator Pelayanan Rumah Sakit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPendapatan 
         Caption         =   "Laporan Pendapatan Ruangan"
      End
   End
   Begin VB.Menu mnuw 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuc 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mbantuan 
      Caption         =   "Ban&tuan"
      Begin VB.Menu mTentang 
         Caption         =   "Medifirst2000"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "MDIUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim sepuh As Boolean

Private Sub BSDJP_Click()
    strCetak = "LapKunjunganSt_PnyktPsn"
    frmLapRKP_KPSK.Show
End Sub

Private Sub BSDKP_Click()
    strCetak = "LapKunjunganKonPulang_Status"
    frmLapRKP_KPSK.Show
End Sub

Private Sub BSJ_Click()
    strCetak = "LapKunjunganJenisStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub LapSensusHarian_Click()
    'frmLapSensusHarian.Show
    FrmlaporanSensusharian.Show
End Sub



Private Sub MCetakFormulirStok_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MCetakFormulirStokNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDaftarCetakInputStokOpnameNM.Show
End Sub

Private Sub MDaftarDokumenRekamMedis_Click()
    frmDaftarDokumenRekamMedisPasien.Show
End Sub

'Private Sub MDIForm_Load()
'
'    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
'    Set rs = Nothing
'
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'    strNmPegawai = rs.Fields("NamaLengkap").Value
'    Set rs = Nothing
'
'    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
'    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
'    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
'    mnlogout.Caption = "Log Off..." & strNmPegawai
'
''    strSQL = "Select StatusFIFO From SettingDataUmum"
''    Call msubRecFO(dbRst, strSQL)
''    If dbRst.EOF = True Then
''        bolStatusFIFO = False
''    Else
''        If dbRst("StatusFIFO") = 0 Then
''            bolStatusFIFO = False
''        Else
''            bolStatusFIFO = True
''        End If
''    End If
'
'    strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
'    Call msubRecFO(dbRst, strSQL)
'    If dbRst.EOF = True Then
'        bolStatusFIFO = False
'    Else
'        If dbRst("MetodeStokBarang") = 0 Then
'            bolStatusFIFO = False
'        Else
'            bolStatusFIFO = True
'        End If
'    End If
'    Unload frmDaftarPasienRI
'
'End Sub

Private Sub MDIForm_Load()
    On Error GoTo errLoad
    Call openConnection
    strSQL = "SELECT * FROM DataPegawai WHERE IdPegawai = '" & strIDPegawaiAktif & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    strNmPegawai = rs.Fields("NamaLengkap").Value
    Set rs = Nothing
    StatusBar1.Panels(1).Text = "Nama User : " & strNmPegawai
    StatusBar1.Panels(2).Text = "Nama Ruangan : " & mstrNamaRuangan
    StatusBar1.Panels(5).Text = "Nama Komputer : " & strNamaHostLocal
    StatusBar1.Panels(6).Text = "Server : " & strServerName & " (" & strDatabaseName & ")"
    mnlogout.Caption = "Log Off..." & strNmPegawai
    '----
    frmPemakaianObatAlkes.NotUseRacikan = True
    '----
'    If mblnAdmin = False Then
'        MTransaksi.Enabled = False
'    Else
'        MTransaksi.Enabled = True
'    End If

    strSQL = "SELECT TerminBayarFakturSupplier, PersentasePpn, PersentaseLimitDiscount, PersentaseJasaPenulisResep, BiayaAdministrasi " & _
    " From SettingDataPendukung" & _
    " WHERE (KdInstalasi = '" & mstrKdInstalasiLogin & "')"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
        typSettingDataPendukung.intTerminBayarFaktur = 0
        typSettingDataPendukung.realJasaPenulisResep = 0
        typSettingDataPendukung.realLimitDiscount = 0
        typSettingDataPendukung.realPPn = 0
        typSettingDataPendukung.curBiayaAdministrasi = 0
    Else
        typSettingDataPendukung.intTerminBayarFaktur = rs("TerminBayarFakturSupplier").Value
        typSettingDataPendukung.realJasaPenulisResep = rs("PersentaseJasaPenulisResep").Value
        typSettingDataPendukung.realLimitDiscount = rs("PersentaseLimitDiscount").Value
        typSettingDataPendukung.realPPn = rs("PersentasePpn").Value
        typSettingDataPendukung.curBiayaAdministrasi = rs("BiayaAdministrasi").Value
    End If

    strSQL = "SELECT JmlPembulatanHarga, JumlahBAdminOAPerBaris, JumlahBAdminTMPerHari FROM MasterDataPendukung"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJmlPembulatanHarga = 0
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
        typSettingDataPendukung.intJumlahBAdminTMPerHari = 0
    Else
        typSettingDataPendukung.intJmlPembulatanHarga = dbRst(0)
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(1)
        typSettingDataPendukung.intJumlahBAdminTMPerHari = dbRst(2)
    End If
    
    strSQL = "SELECT JmlBarisOAPerTarifAdminOA from SettingBiayaAdministrasi"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = 0
    Else
        
        typSettingDataPendukung.intJumlahBAdminOAPerBaris = dbRst(0)
        
    End If

    strSQL = "Select MetodeStokBarang From SuratKeputusanRuleRS where statusenabled=1"
    Call msubRecFO(dbRst, strSQL)
    If dbRst.EOF = True Then
        bolStatusFIFO = False
    Else
        If dbRst("MetodeStokBarang") = 0 Then
            bolStatusFIFO = False
        Else
            bolStatusFIFO = True
        End If
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub


Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then Exit Sub
    PopupMenu mnudata
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim q As String

    If sepuh = True Then
        q = MsgBox("Log Off user " & strNmPegawai & " ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then
            Unload frmLogin
            Cancel = 1
        Else
            Cancel = 0
            frmLogin.Show
        End If
        sepuh = False
    Else
        q = MsgBox("Tutup aplikasi ", vbQuestion + vbOKCancel, "Konfirmasi")
        If q = 2 Then

            Unload frmLogin
            Cancel = 1
        Else
            dTglLogout = Now
            Call subSp_HistoryLoginAplikasi("U")
            Cancel = 0
        End If
    End If

End Sub

Private Sub mGantiKataKunci_Click()
    frmLoginEditAccount.Show
End Sub

Private Sub MInformasiPemakaianBarang_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarang_Click()
    mstrKdKelompokBarang = "02"  'medis
    frmInfoPesanBarang.Show
End Sub

Private Sub MInformasiPemesananPenerimaanBarangNonMedis_Click()
    mstrKdKelompokBarang = "01"  'non medis
    frmInfoPesanBarangNM.Show
End Sub

Private Sub MInformasiTarifPelayanan_Click()
    frmInformasiTarifPelayanan.Show
End Sub

Private Sub mnCetakLembarInput_Click()
    frmDaftarCetakInputStokOpname.Show
End Sub

Private Sub MKondisiBarangNonMedis_Click()
    frmKondisiBarangNM.Show
End Sub

Private Sub MLaporanSaldoBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmLaporanSaldoBarangMedis_v3.Show
End Sub

Private Sub MLaporanSaldoBarangNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmLaporanSaldoBarangNM_v3.Show
End Sub

Private Sub MMutasiBarangNonMedis_Click()
    frmMutasiBarangNM.Show
End Sub

Private Sub MNilaiPersediaan_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmNilaiPersediaan.Show
End Sub

Private Sub mnInputStokOpname_Click()
    frmStokOpname.Show
End Sub

Private Sub MNilaiPersediaanNonMedis_Click()

    mstrKdKelompokBarang = "01"
    frmNilaiPersediaanNM.Show

End Sub

Private Sub mnlogout_Click()

    Dim adoCommand As New ADODB.Command

    openConnection
    sepuh = True
    strQuery = "UPDATE Login SET Status = '0' " & _
    "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
    adoCommand.ActiveConnection = dbConn
    adoCommand.CommandText = strQuery
    adoCommand.CommandType = adCmdText
    adoCommand.Execute

    dTglLogout = Now
    Call subSp_HistoryLoginAplikasi("U")

    Unload Me

End Sub

Private Sub mnLPBPP_Click()
    FrmBukuRegisterPelayanan.Show
End Sub

Private Sub mnMonitoring_Click()
    frmMonitoringPembayaran.Show
End Sub

Private Sub mnSelesai_Click()

    Dim pesan As VbMsgBoxResult
    Dim adoCommand As New ADODB.Command
    pesan = MsgBox("Tutup aplikasi ", vbQuestion + vbYesNo, "Konfirmasi")
    If pesan = vbYes Then

        openConnection
        strQuery = "UPDATE Login SET Status = '0' " & _
        "WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        adoCommand.ActiveConnection = dbConn
        adoCommand.CommandText = strQuery
        adoCommand.CommandType = adCmdText
        adoCommand.Execute

        dTglLogout = Now
        Call subSp_HistoryLoginAplikasi("U")
        End
    Else
    End If

End Sub

Private Sub mnuai_Click()
    MDIUtama.Arrange vbArrangeIcons
End Sub

Private Sub mnuBDataDiag_Click()
    frmDataDiagnosa.Show
End Sub

Private Sub mnuBDPOAK_Click()
    frmDaftarPakaiAlkesKaryawan.Show
End Sub

Private Sub mnubrp_Click()
    FrmBukuRegisterPasien.Show
    'FrmBukuRegister.Show
End Sub

Private Sub mnuc_Click()
    MDIUtama.Arrange vbCascade
End Sub

Private Sub mnucdp_Click()
    frmDaftarPasienRI.Show
End Sub

Private Sub mnuClosingDataPelayananTMOAApotik_Click()
    frmClosingDataPelayananTM_OA_Apotik.Show
End Sub

Private Sub mnuDaftarBayiLahir_Click()
    frmDaftarBayiLahir.Show
End Sub

Private Sub mnuDaftarPengirimanDarah_Click()
    frmDaftarPengirimanDarah.Show
End Sub

Private Sub mnuDetailDiagnosaKeperawatan_Click()
    frmDetailDiagnosaKeperawatan.Show
End Sub

Private Sub mnuDiagnosa_Click()
    frmMasterDiagnosaAskep.Show
End Sub

Private Sub mnudpl_Click()
    frmDaftarPasienLama.Show
End Sub

Private Sub mnudpm_Click()
    frmDaftarPasienMeninggal.Show
End Sub

Private Sub mnudsmp_Click()
    frmLapMorbiditas.Show
    frmLapMorbiditas.Caption = "Medifirst2000 - Data Keadaan Morbiditas Pasien"
End Sub

Private Sub mnuipb_Click()
    frmInfoPesanBarang.Show
End Sub

Private Sub mnuipoa_Click()
    frmDaftarPakaiAlkes.Show
End Sub

Private Sub mnuiprs_Click()
    frmLapIndPlynRS.Show
    frmLapIndPlynRS.Caption = "Medifirst2000 - Indikator Pelayanan RS"
End Sub

Private Sub mnuJenisPersalinanEventBayi_Click()
    frmMasterVKbersalin.Show
End Sub

Private Sub mnuLap1_Click()
    frmLapSensusHarian2.Show
End Sub

Private Sub mnuLap2_Click()
    frmLapSensusHarian3.Show
End Sub

Private Sub mnuLap3_Click()
    frmLaporanPendapatanRuangan.Show
End Sub

Private Sub mnuLapPendapatan_Click()
frmDaftarPendapatanRuangan.Show
End Sub

Private Sub mnuMastDiagnosaKeperawatan_Click()
    frmMasterDiagnosaKeperawatan.Show
End Sub

Private Sub mnuPakaiBahan_Click()
    frmPemakaianBahanAlat.Show
End Sub

Private Sub mnupb_Click()
    frmPemesananBarang.Show
End Sub

Private Sub mnuPesanPelayananTMOA_Click()
    frmInfoPesanPelayananTMOA.Show
End Sub

Private Sub mnupp_Click()
    frmPaketLayanan.Show
End Sub

Private Sub mnur10bp_Click()
    FrmPeriodeLaporanTopTen.Show
    FrmPeriodeLaporanTopTen.Caption = "Medifirst2000 - Rekapitulasi 10 Besar Penyakit"
End Sub

Private Sub mnurkri_Click()
    frmLapRkpKmrRI.Show
    frmLapRkpKmrRI.Caption = "Medifirst2000 - Rekapitulasi Kamar Rawat Inap"
End Sub

Private Sub mnusb_Click()
    frmStokBrg.Show
End Sub

Private Sub mnuTujuanNRencanaTindakan_Click()
    frmTujuanNRencanaTindakan.Show
End Sub

Private Sub MPenyebabDiagosaKeperawatan_Click()
    frmPenyebabDiagnosaKeperawatan.Show
End Sub

Private Sub MRekapitulasiTransaksiBarang_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmDataTransaksiBarangNM.Show
End Sub

Private Sub MRekapitulasiTransaksiBarangNonMedis_Click()
    mstrKdKelompokBarang = "01" 'non medis
    frmDataTransaksiBarangNM.Show
End Sub

Private Sub mSettingPrinter_Click()
    frmSetupPrinter2.Show
End Sub

Private Sub MStokBarangNonMedis_Click()
    frmStokBarangNonMedis.Show
End Sub

Private Sub MStokOpname_Click()
    mstrKdKelompokBarang = "02" 'medis
    frmStokOpname.Show
End Sub

Private Sub MStokOpnameNonMedis_Click()
    mstrKdKelompokBarang = "01"
    frmStokOpnameNM.Show
End Sub

Private Sub mTentang_Click()
    frmAbout.Show
End Sub

Private Sub RKPR_Click()
    strCetak = "LapKunjunganRujukanBStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBD_Click()
    strCetak = "LapKunjunganBDiagnosa"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBKel_Click()
    strCetak = "LapKunjunganKelasStatus"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBW_Click()
    strCetak = "LapKunjunganBwilayah"
    frmLapRKP_KPSK.Show
End Sub

Private Sub RPBWD_Click()
    strCetak = "LapKunjunganPasienBDiagnosaWilayah"
    frmLapRKP_KPSK.Show
End Sub

