VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCtkIndikatorPelPerRuang 
   Caption         =   "Medifirst2000 - Indikator Pelayanan Per Ruangan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCtkIndikatorPelPerRuang.frx":0000
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
Attribute VB_Name = "frmCtkIndikatorPelPerRuang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReportBulan As New crIndikatorPelPerRuang

Private Sub Form_Load()
    On Error GoTo errLoad

    Dim tanggal As String
    Dim laporan As String

    Set dbcmd = New ADODB.Command

    tanggal = Format(frmIndikatorPelPerRuang.dtpAwal.Value, "yyyy")

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2

    dbcmd.ActiveConnection = dbConn
    dbcmd.CommandText = "select * from IndikatorPelPerRuangX where kode = '" & varNoCetak & "'"

    With ReportBulan
        .Database.AddADOCommand dbConn, dbcmd

        .txtNamaRS.SetText strNNamaRS & " " & strKelasRS & " " & strKetKelasRS
        .txtAlamat.SetText "KABUPATEN " & strNKotaRS
        .txtAlamat2.SetText strNAlamatRS & " " & "Telp." & " " & strNTeleponRS
        .txtHari.SetText varHari

        If varStatusBulanan = False Then
            .txtPeriode.SetText tanggal
            .txtTahunBulan.SetText "TAHUN"
            .txtIsiIndikator.SetText "Tahun " & tanggal & " secara umum berada pada parameter XXXXX, dengan indikator sbb :"
        Else
            .txtPeriode.SetText MonthName(Int(varBulan)) + " " + tanggal
            .txtTahunBulan.SetText "BULAN"
            .txtIsiIndikator.SetText "Tahun " & tanggal + " " + MonthName(Int(varBulan)) & " secara umum berada pada parameter XXXXX, dengan indikator sbb :"
        End If
        .txtNamaUser.SetText "Operator " + strNmPegawai

        .strRuang.SetUnboundFieldSource ("{ado.NamaRuangan}")
        .unTT.SetUnboundFieldSource ("{ado.TT}")
        .unPMasuk.SetUnboundFieldSource ("{ado.Masuk}")
        .unKeluarHidup.SetUnboundFieldSource ("{ado.KeluarH}")

        .unLamaRawat.SetUnboundFieldSource ("{ado.LamaRawat}")
        .unHariPRawat.SetUnboundFieldSource ("{ado.HP}")

        .unMatiK.SetUnboundFieldSource ("{ado.KeluarM1}")
        .unMatiL.SetUnboundFieldSource ("{ado.KeluarM2}")
        .unHari.SetUnboundFieldSource ("{ado.Periode}")

    End With
    CRViewer1.ReportSource = ReportBulan

    If vLaporan = "Print" Then

        ReportBulan.PrintOut False
        Unload Me
    Else
        With CRViewer1
            .Zoom 1
            .ViewReport

        End With
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
errLoad:
    Screen.MousePointer = vbDefault
    msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCtkIndikatorPelPerRuang = Nothing
End Sub
