VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmcetakviewer 
   Caption         =   "Medifirst2000 - Cetak"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmCetakViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmcetakviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crRkpPsnJenisPeriksa
Dim rpttahun As New CrTndakandanPeriksaPTahun

Private Sub Form_Load()
    Select Case strCetak2
        Case "LapKunjunganJenisPeriksahari"
            Select Case strCetak3
                Case "JenisPeriksaBInstalasiAsal"
                    Call LapKunjunganJenisPeriksaHariperAsalInstalasi
                Case "JenisPeriksaBJenispasien"
                    Call lapkunjunganperhariperjenispasien
            End Select
        Case "LapKunjunganJenisPeriksabulan"
            Select Case strCetak3
                Case "JenisPeriksaBInstalasiAsal"
                    Call lakunjunganbulaninstalasiasal
                Case "JenisPeriksaBJenispasien"
                    Call JenisPeriksaBJenispasienBulan
            End Select
        Case "LapKunjunganJenisPeriksaTahun"
            Select Case strCetak3
                Case "JenisPeriksaBInstalasiAsal"
                    Call JenisPeriksaBInstalasiAsalTahun
                Case "JenisPeriksaBJenispasien"
                    Call JenisPeriksaBJenispasienTahun
            End Select
            '========================================================
        Case "LapKunjunganJenisTindakanHari"
            Select Case strCetak3
                Case "LapKunjunganJenisTindakanBinstalasiAsal"
                    Call LapKunjunganJenisTindakanBinstalasiAsal
                Case "LapKunjunganJenisTindakanBJenisPasienHari"
                    Call LapKunjunganJenisTindakanBJenisPasienHari
            End Select
        Case "LapKunjunganJenisTindakanBulan"
            Select Case strCetak3
                Case "LapKunjunganJenisTindakanBinstalasiAsalBulan"
                    Call LapKunjunganJenisTindakanBinstalasiAsalBulan
                Case "LapKunjunganJenisTindakanBJenisPasienBulan"
                    Call LapKunjunganJenisTindakanBJenisPasienBulan
            End Select
        Case "LapKunjunganJenisTindakantahun"
            Select Case strCetak3
                Case "LapKunjunganJenisTindakanBinstalasiAsaltahun"
                    Call LapKunjunganJenisTindakanBinstalasiAsaltahun
                Case "LapKunjunganJenisTindakanBJenisPtahun"
                    Call LapKunjunganJenisTindakanBJenisPtahun
            End Select

    End Select
End Sub

Private Sub JenisPeriksaBInstalasiAsalTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usjenispelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usinstalasiasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtpilihan.SetText ("Jenis Pasien")
        .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
    Set frmcetakviewer = Nothing
End Sub

Private Sub JenisPeriksaBJenispasienTahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usjenispelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usinstalasiasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtpilihan.SetText ("Jenis Pasien")
        .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
    Set frmcetakviewer = Nothing
End Sub

Private Sub JenisPeriksaBJenispasienBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText "LAPORAN KUNJUNGAN PASIEN BERDASARKAN JENIS PERIKSA"
        .txtTanggal.SetText CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy"))
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
        .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Jenis Pasien")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisTindakanBJenisPtahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usjenispelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usinstalasiasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtpilihan.SetText ("Jenis Pasien")
        .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If

        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
    Set frmcetakviewer = Nothing
End Sub

Private Sub LapKunjunganJenisTindakanBinstalasiAsaltahun()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With rpttahun
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .Priode.SetText ("Periode  : " & CStr(Format(mdTglAwal, "yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .Udate.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usjenispelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .usinstalasiasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtpilihan.SetText ("Instalasi Asal")
        .UJml.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport rpttahun, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = rpttahun
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
    Set frmcetakviewer = Nothing
End Sub

Private Sub LapKunjunganJenisTindakanBJenisPasienBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Jenis pasien")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
    Set frmcetakviewer = Nothing
End Sub

Private Sub LapKunjunganJenisTindakanBJenisPasienHari()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "mm-dd-yy")) & " s/d " & CStr(Format(mdTglAkhir, "mm-dd-yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Jenis Pasien")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisTindakanBinstalasiAsalBulan()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy")))

        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Instalasi Asal")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisTindakanBinstalasiAsal()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .txtJudul.SetText ("LAPORAN  KUNJUNGAN PASIEN BERDASARKAN JENIS TINDAKAN ")
        .txtTanggal.SetText ("Periode  : " & CStr(Format(mdTglAwal, "mm-dd-yy")) & " s/d " & CStr(Format(mdTglAkhir, "mm-dd-yyyy")))
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPelayanan}")
        .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Instalasi Asal")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub lakunjunganbulaninstalasiasal()
    Call openConnection
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText "LAPORAN KUNJUNGAN PASIEN BERDASARKAN BERDASARKAN JENIS PERIKSA"
        .txtTanggal.SetText CStr(Format(mdTglAwal, "MMMM yyyy")) & " s/d " & CStr(Format(mdTglAkhir, "MMMM yyyy"))
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
        .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Instalasi Asal")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub LapKunjunganJenisPeriksaHariperAsalInstalasi()
    Call openConnection
    Set frmcetakviewer = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText "LAPORAN KUNJUNGAN PASIEN BERDASARKAN JENIS PERIKSA"
        .txtTanggal.SetText CStr(Format(mdTglAwal, "mm-dd-yy")) & " s/d " & CStr(Format(mdTglAkhir, "mm-dd-yyyy"))
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
        .UsInstasal.SetUnboundFieldSource ("{ado.InstalasiAsal}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Instalasi Asal")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub

Private Sub lapkunjunganperhariperjenispasien()

    Call openConnection
    Set frmcetakviewer = Nothing
    Dim adocomd As New ADODB.Command
    Set adocomd = Nothing
    Me.WindowState = 2
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = strSQL
    With Report
        .Database.AddADOCommand dbConn, adocomd
        .txtJudul.SetText "LAPORAN KUNJUNGAN PASIEN BEDASARKAN JENIS PERIKSA"
        .txtTanggal.SetText CStr(Format(mdTglAwal, "mm-dd-yy")) & " s/d " & CStr(Format(mdTglAkhir, "mm-dd-yyyy"))
        .txtNamaRS.SetText strNNamaRS
        .txtAlamat.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
        .txtAlamat2.SetText strEmail
        .UdPeriode.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usRuangan.SetUnboundFieldSource ("{ado.RuanganPelayanan}")
        .UsJnsPelayanan.SetUnboundFieldSource ("{ado.JenisPeriksa}")
        .UsInstasal.SetUnboundFieldSource ("{ado.JenisPasien}")
        .usJK.SetUnboundFieldSource ("{ado.JK}")
        .txtjenis.SetText ("Jenis Pasien")
        .JMlPasien.SetUnboundFieldSource ("{ado.JmlPelayanan}")
        If sUkuranKertas = "" Then
            sUkuranKertas = "5"
            sOrientasKertas = "2"
            sDuplex = "0"
        End If
        settingreport Report, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
    End With

    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom (75)
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcetakviewer = Nothing
End Sub
