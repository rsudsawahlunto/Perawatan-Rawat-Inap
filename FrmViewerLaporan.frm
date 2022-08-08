VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmViewerLaporan 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FrmViewerLaporan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "FrmViewerLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportBukuBesar As New crBukuBesar
Dim reportTopten As New crDiagnosaTopTen2
Dim reporrtoptengrafik As New crDiagnosaTopTenGrafik

Private Sub Form_Load()

    Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass

    Call openConnection
    Me.WindowState = 2
    Dim tanggal As String

    Select Case cetak
        Case "BukuBesar"

            adocomd.ActiveConnection = dbConn
            adocomd.CommandText = "SELECT * FROM V_BukuRegisterPasienRI1 " _
            & "WHERE (TglMasuk BETWEEN '" _
            & Format(FrmBukuRegister.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmBukuRegister.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "')" _
            & " AND kdruangan = '" & mstrKdRuangan & "'" & _
            "ORDER BY NamaRuangan,JenisPasien"
            adocomd.CommandType = adCmdUnknown
            reportBukuBesar.Database.AddADOCommand dbConn, adocomd

            With reportBukuBesar
                .Text19.SetText strNNamaRS
                .Text20.SetText strNAlamatRS
                .Text21.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .txtRuang.SetText strNNamaRuangan
                .txttgl.SetText Format(FrmBukuRegister.DTPickerAwal, "dd/MM/yyyy") & "  s/d  " & Format(FrmBukuRegister.DTPickerAkhir, "dd/MM/yyyy")
                .udtTglMasuk.SetUnboundFieldSource "{ado.tglmasuk}"
                .usCM.SetUnboundFieldSource "{ado.nocm}"
                .usPasien.SetUnboundFieldSource "{ado.namapasien}"
                .usAlamat.SetUnboundFieldSource "{ado.alamat}"
                .usPekerjaan.SetUnboundFieldSource "{ado.pekerjaan}"
                .usUmur.SetUnboundFieldSource "{ado.umur}"
                .usJK.SetUnboundFieldSource "{ado.jk}"
                .usStatus.SetUnboundFieldSource "{ado.status}"
                .usRujukan.SetUnboundFieldSource "{ado.AsalRujukan}"
                .usCrMsk.SetUnboundFieldSource "{ado.CaraMasuk}"
                .usKelas.SetUnboundFieldSource "{ado.Kelas}"
                .usKlpkPasien.SetUnboundFieldSource "{ado.jenispasien}"
                .usSMF.SetUnboundFieldSource "{ado.SMF}"
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reportBukuBesar, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
            End With
            CRViewer1.ReportSource = reportBukuBesar

            Screen.MousePointer = vbHourglass

            If vLaporan = "Print" Then
                reportBukuBesar.PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ViewReport
                    .Zoom 100
                End With
            End If
            'Rekapitulasi 10 besar Penyakit
        Case "RekapTopten"

            adocomd.CommandText = "sELECT * FROM V_RekapitulasiDiagnosaTopTen " _
            & "WHERE (TglPeriksa BETWEEN '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "') " _
            & " and NamaRuangan = '" & mstrNamaRuangan & "' ORDER BY instalasi,diagnosa"

            adocomd.CommandType = adCmdText
            reportTopten.Database.AddADOCommand dbConn, adocomd

            If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
                tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy")
            Else
                tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            End If

            With reportTopten
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS
                .Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .txtPeriode2.SetText tanggal
                .txtinstalasi.SetText ""
                .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
                .UsDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
                .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reportTopten, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
            End With

            CRViewer1.ReportSource = reportTopten

            Screen.MousePointer = vbHourglass

            If vLaporan = "Print" Then
                reportTopten.PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ViewReport
                    .Zoom 100
                End With
            End If

        Case "RekapToptenGrafik"

            adocomd.CommandText = "sELECT * FROM V_RekapitulasiDiagnosaTopTen " _
            & "WHERE (TglPeriksa BETWEEN '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "')  " _
            & " and NamaRuangan = '" & mstrNamaRuangan & "' ORDER BY instalasi,diagnosa"

            adocomd.CommandType = adCmdText
            reporrtoptengrafik.Database.AddADOCommand dbConn, adocomd

            If Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") = Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy") Then
                tanggal = "Tanggal Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") '& " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            Else
                tanggal = "Periode Kunjungan  : " & " " & Format(FrmPeriodeLaporanTopTen.DTPickerAwal, "dd MMMM yyyy") & " S/d " & Format(FrmPeriodeLaporanTopTen.DTPickerAkhir, "dd MMMM yyyy")
            End If

            With reporrtoptengrafik
                .Text1.SetText strNNamaRS
                .Text2.SetText strNAlamatRS
                .Text3.SetText strNKotaRS & " " & "Kode Pos " & " " & strNKodepos & " " & "Telp." & " " & strNTeleponRS
                .txtPeriode2.SetText tanggal
                .txtinstalasi.SetText ""
                .usSMF.SetUnboundFieldSource ("{ado.instalasi}")
                .UsDiagnosa.SetUnboundFieldSource ("{ado.diagnosa}")
                .unJumlahPasien.SetUnboundFieldSource ("{ado.jumlahpasien}")
                .SelectPrinter sDriver, sPrinter, vbNull
                settingreport reporrtoptengrafik, sPrinter, sDriver, sUkuranKertas, sDuplex, sOrientasKertas
            End With

            CRViewer1.ReportSource = reporrtoptengrafik

            Screen.MousePointer = vbHourglass

            If vLaporan = "Print" Then
                reporrtoptengrafik.PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ViewReport
                    .Zoom 100
                End With
            End If

    End Select

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()

    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmViewerLaporan = Nothing
End Sub

