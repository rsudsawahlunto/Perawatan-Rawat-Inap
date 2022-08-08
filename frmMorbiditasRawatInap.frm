VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMorbiditasRI 
   Caption         =   "Morbiditas "
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMorbiditasRawatInap.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   5895
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
Attribute VB_Name = "frmMorbiditasRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report As New crMorbiditasRawatInap
Dim adoCommand As New ADODB.Command
'Dim strSQL As String
Dim tanggal As String

Private Sub Form_Load()
    openConnection
    Set adoCommand.ActiveConnection = dbConn
    Set frmMorbiditasRI = Nothing

    If Format(mdTglAwal, "dd MMMM yyyy") = Format(mdTglAkhir, "dd MMMM yyyy") Then
        tanggal = "Tanggal Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") '& " S/d " & Format(mdtglAkhir, "dd MMMM yyyy")
    Else
        tanggal = "Periode Kunjungan  : " & " " & Format(mdTglAwal, "dd MMMM yyyy") & " S/d " & Format(mdTglAkhir, "dd MMMM yyyy")
    End If

    adoCommand.CommandText = strSQL
    adoCommand.CommandType = adCmdText
    With Report
        .Database.AddADOCommand dbConn, adoCommand
        .txtJudul.SetText "DATA KEADAAN MORBIDITAS RAWAT INAP SURVEILANS TERPADU RUMAH SAKIT"
        .txtJudul2.SetText "FORMULIR RL 2a1"
        .txtPeriode.SetText tanggal
        .usNoDTD.SetUnboundFieldSource ("{ado.NoDTD}")
        .usNoDT.SetUnboundFieldSource ("{ado.NoDTerperinci}")
        .usNamaDTD.SetUnboundFieldSource ("{ado.NamaDTD}")
        .unKel1.SetUnboundFieldSource ("{ado.Kel_Umur1}")
        .unKel2.SetUnboundFieldSource ("{ado.Kel_Umur2}")
        .unKel3.SetUnboundFieldSource ("{ado.Kel_Umur3}")
        .unKel4.SetUnboundFieldSource ("{ado.Kel_Umur4}")
        .unKel5.SetUnboundFieldSource ("{ado.Kel_Umur5}")
        .unKel6.SetUnboundFieldSource ("{ado.Kel_Umur6}")
        .unKel7.SetUnboundFieldSource ("{ado.Kel_Umur7}")
        .unKel8.SetUnboundFieldSource ("{ado.Kel_Umur8}")
        .unKelL.SetUnboundFieldSource ("{ado.Kel_L}")
        .unKelP.SetUnboundFieldSource ("{ado.Kel_P}")
        .unKelH.SetUnboundFieldSource ("{ado.Kel_H}")
        .unKelM.SetUnboundFieldSource ("{ado.Kel_M}")
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS
        .Text3.SetText strNKotaRS & " " & strNKodepos & " Telp. " & strNTeleponRS
        .SelectPrinter sDriver, sPrinter, vbNull
        settingreport Report, sPrinter, sDriver, crPaperLegal, sDuplex, crLandscape
    End With
    '\---------------------------------------------------/
    Screen.MousePointer = vbHourglass

    If vLaporan = "Print" Then
        Report.PrintOut False
        Unload Me
    Else

        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 1
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMorbiditasRI = Nothing
End Sub
