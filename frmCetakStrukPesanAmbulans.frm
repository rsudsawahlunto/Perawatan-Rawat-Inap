VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakStrukPesanAmbulans 
   Caption         =   "Medifrst2000 - Struk Pemesanan dari Ruangan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   Icon            =   "frmCetakStrukPesanAmbulans.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6930
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7005
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   0   'False
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
Attribute VB_Name = "frmCetakStrukPesanAmbulans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crStrukPesanAmbulans

Private Sub Form_Load()
On Error GoTo errLoad
Dim adocomd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    Me.WindowState = 2
    Call openConnection
    
    adocomd.ActiveConnection = dbConn
    adocomd.CommandText = "SELECT NamaPasien,NamaPJawabKeluarga,NoTlpHP,TujuanOrder,NamaPelayanan,TglPelayanan,NamaTempatTujuan,AlamatLengkapTempatTujuan,QtyPelayanan FROM  V_DaftarPesanAmbulans WHERE NoOrder = '" & NoOrder & "'"
    adocomd.CommandType = adCmdText
    Report.Database.AddADOCommand dbConn, adocomd

    With Report
        .Text1.SetText strNNamaRS
        .Text2.SetText strNAlamatRS & ", " & strNKotaRS & " - " & strNKodepos & " Telp. " & strNTeleponRS
        .Text3.SetText strWebsite & ", " & strEmail
        
'        Call msubRecFO(dbRst, "SELECT NoOrder, TglOrder, RuanganPemesan, UserPemesan FROM  V_StrukOrderInformasiNonMedisRuangan WHERE NoOrder = '" & mstrNoOrder & "'")
        Call msubRecFO(dbRst, "SELECT * from StrukOrder WHERE NoOrder = '" & NoOrder & "'")
        If rs.EOF = False Then
            .txtNoOrder.SetText IIf(IsNull(dbRst("NoOrder")), "", dbRst("NoOrder"))
            .txtTglOrder.SetText IIf(IsNull(dbRst("TglOrder")), "", dbRst("TglOrder"))
'            .txtRuanganPemesan.SetText IIf(IsNull(dbRst("RuanganPemesan")), "", dbRst("RuanganPemesan"))
'            .txtUserPemesan.SetText IIf(IsNull(dbRst("UserPemesan")), "", dbRst("NamaPasien"))
        End If
        
'        .usNoOrder.SetUnboundFieldSource ("{ado.NoOrder}")
        .unQty.SetUnboundFieldSource ("{ado.QtyPelayanan}")
        .usNamaPasien.SetUnboundFieldSource ("{ado.NamaPasien}")
        .usNama.SetUnboundFieldSource ("{ado.NamaPJawabKeluarga}")
        .unNoTlpHP.SetUnboundFieldSource ("{ado.NoTlpHP}")
        .usTujuanOrder.SetUnboundFieldSource ("{ado.TujuanOrder}")
        .usNamaPelayanan.SetUnboundFieldSource ("{ado.NamaPelayanan}")
        .udTglPelayanan.SetUnboundFieldSource ("{ado.TglPelayanan}")
        .usTempatTujuan.SetUnboundFieldSource ("{ado.NamaTempatTujuan}")
        .usAlamatTujuan.SetUnboundFieldSource ("{ado.AlamatLengkapTempatTujuan}")
'        .PrintOut False
    End With
    If vLaporan = "view" Then
        Screen.MousePointer = vbHourglass
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom (100)
        End With
    Else
        Report.PrintOut False
        Unload Me
    End If
    Screen.MousePointer = vbDefault
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCetakStrukPesanAmbulans = Nothing
End Sub
