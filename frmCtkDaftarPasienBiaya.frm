VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkDaftarPasienBiaya 
   Caption         =   "Cetak Dokumen Rekam Medis Pasien"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmCtkDaftarPasienBiaya.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11100
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCtkDaftarPasienBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crDaftarPasienBiaya

Private Sub Form_Load()
    Dim adocomd As New ADODB.Command
    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection

    Me.WindowState = 2
    Report.txtNamaRS.SetText strNNamaRS
    Report.txtAlamatRS.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & "  Telp. " & strNTeleponRS
    Report.txtWebsiteRS.SetText strWebsite & ", " & strEmail
    Report.Text14.SetText "RSUD " & strNKotaRS

    If frmDaftarPasienRI.optPasAktif.Value = True Then
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "SELECT  * From V_DaftarPasienRIAktif where Ruangan='" & strNNamaRuangan & "' and ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%')  AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'"
        adocomd.CommandType = adCmdUnknown
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .usNoRegister.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
            .usJK.SetUnboundFieldSource ("{ado.JK}")
            .usUmur.SetUnboundFieldSource ("{ado.Umur}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .udTglMasuk.SetUnboundFieldSource ("{ado.TglMasuk}")
            .usNoKamar.SetUnboundFieldSource ("{ado.NoKamar}")
            .usNoBed.SetUnboundFieldSource ("{ado.NoBed}")
        End With
        Report.txtDaftarKamar.SetText (UCase(frmDaftarPasienRI.optPasAktif.Caption))
    ElseIf frmDaftarPasienRI.optPasNonAktif.Value = True Then
        adocomd.ActiveConnection = dbConn
        adocomd.CommandText = "SElect * from V_DaftarPasienRIPindahKamar where ([Nama Pasien] like '%" & frmDaftarPasienRI.txtParameter.Text & "%' or NoCM like '%" & frmDaftarPasienRI.txtParameter.Text & "%') AND KdRuanganTujuan='" & mstrKdRuangan & "' AND JenisPasien LIKE '%" & frmDaftarPasienRI.dcJenisPasien.Text & "%' AND Kelas LIKE '%" & frmDaftarPasienRI.dcKelas.Text & "%'"
        adocomd.CommandType = adCmdUnknown
        Report.Database.AddADOCommand dbConn, adocomd
        With Report
            .usNoRegister.SetUnboundFieldSource ("{ado.NoPendaftaran}")
            .usNoCM.SetUnboundFieldSource ("{ado.NoCM}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.Nama Pasien}")
            .usJK.SetUnboundFieldSource ("{ado.JK}")
            .usUmur.SetUnboundFieldSource ("{ado.Umur}")
            .usKelas.SetUnboundFieldSource ("{ado.Kelas}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.JenisPasien}")
            .Text10.SetText "Tanggal Pindah"
            .Text1.SetText "Ruang Asal"
            .Text3.SetText "Ruang Tujuan"
            .udTglMasuk.SetUnboundFieldSource ("{ado.TglPindah}")
            .usNoKamar.SetUnboundFieldSource ("{ado.Ruangan Asal}")
            .usNoBed.SetUnboundFieldSource ("{ado.Ruangan Tujuan}")
        End With
        Report.txtDaftarKamar.SetText (UCase(frmDaftarPasienRI.optPasNonAktif.Caption))
    End If
    Report.txtTanggal.SetText ("")
    Screen.MousePointer = vbHourglass
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 1
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCtkDaftarPasien = Nothing
End Sub
