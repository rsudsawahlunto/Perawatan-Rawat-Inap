VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmValidasiData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 -Validasi Data"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValidasiData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11085
   Begin MSComctlLib.ProgressBar pbData 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "&Perbaiki Data"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdValidasiData 
      Caption         =   "&Validasi Data"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   6480
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
      Picture         =   "frmValidasiData.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9240
      Picture         =   "frmValidasiData.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmValidasiData.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmValidasiData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim substrNomorRetur As String
Dim subbolSimpan As Boolean
Dim sqlQuery As String
Dim rsQuery As New ADODB.recordset
Dim i As Integer
Dim j As Integer
Dim strStatusData As String

Private Sub subLoadData(Optional s_Kriteria As String)

    On Error GoTo errLoad

    Dim strStatusData As String
    Dim sqlQuery As String
    Dim rsQuery As New ADODB.recordset

    Call subSetGrid
    strSQL = " SELECT NoPendaftaran, RuanganPelayanan, NamaPelayanan, TglPelayanan, Tarif, TarifCito, JmlHutangPenjamin, " & _
    " JmlTanggunganRS , JmlPembebasan , KdPelayananRS, KdRuangan,KdInstalasi,  KdKelas, KdJenisTarif, JmlPelayanan, StatusCITO,IdPegawai " & _
    " From V_DetailBiayaPelayanan4Validasi where NoPendaftaran ='" & mstrNoPen & "'"
    Call msubRecFO(rs, strSQL)

    If rs.EOF = True Then Exit Sub

    For i = 1 To rs.RecordCount
        With fgData
            .TextMatrix(i, 0) = rs("NoPendaftaran")
            .TextMatrix(i, 1) = rs("RuanganPelayanan")
            .TextMatrix(i, 2) = rs("NamaPelayanan")
            .TextMatrix(i, 3) = rs("TglPelayanan")

            .TextMatrix(i, 5) = rs("Tarif")
            .TextMatrix(i, 6) = rs("TarifCito")
            .TextMatrix(i, 7) = rs("JmlHutangPenjamin")
            .TextMatrix(i, 8) = rs("JmlTanggunganRS")
            .TextMatrix(i, 9) = rs("JmlPembebasan")
            .TextMatrix(i, 10) = rs("KdPelayananRS")
            .TextMatrix(i, 11) = rs("KdRuangan")
            .TextMatrix(i, 12) = rs("KdInstalasi")

            .TextMatrix(i, 13) = rs("KdKelas")
            .TextMatrix(i, 14) = rs("KdJenisTarif")
            .TextMatrix(i, 15) = rs("JmlPelayanan")
            .TextMatrix(i, 16) = rs("StatusCITO")
            .TextMatrix(i, 17) = rs("IdPegawai")

            sqlQuery = "SELECT dbo.FB_TakeStatusDataValid('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 11) & "','" & .TextMatrix(i, 10) & "', '" & Format(.TextMatrix(i, 3), "yyyy/mm/dd HH:mm:ss ") & "', " & msubKonversiKomaTitik(.TextMatrix(i, 5)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 6)) & " , " & msubKonversiKomaTitik(.TextMatrix(i, 7)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 8)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 9)) & " )  as StatusData"
            Call msubRecFO(rsQuery, sqlQuery)
            strStatusData = rsQuery.Fields(0).Value
            .TextMatrix(i, 4) = strStatusData

            For j = 0 To 4
                .Col = j
                If .Col = 4 Then
                    If .TextMatrix(i, 4) = "T" Then
                        .Row = i
                        .CellBackColor = vbRed
                        .CellForeColor = vbWhite
                    Else
                        .Row = i
                        .CellBackColor = vbWhite
                        .CellForeColor = vbBlack
                    End If
                End If
            Next j

            .Rows = .Rows + 1
            rs.MoveNext
        End With
    Next i

    Exit Sub
errLoad:
    Call msubPesanError

End Sub

Private Sub subLoadText()

    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 4
            txtIsi.MaxLength = 1
        Case Else
            Exit Sub
    End Select

    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus

End Sub

Private Sub subSetGrid()

    With fgData
        .Rows = 2
        .Cols = 18

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "No Pendaftaran"
        .TextMatrix(0, 1) = "Ruang Pelayanan"
        .TextMatrix(0, 2) = "Nama Pelayanan"
        .TextMatrix(0, 3) = "Tgl Pelayanan"
        .TextMatrix(0, 4) = "Status Data"
        .TextMatrix(0, 5) = "Tarif"
        .TextMatrix(0, 6) = "TarifCito"
        .TextMatrix(0, 7) = "JmlHutangPenjamin"
        .TextMatrix(0, 8) = "JmlTanggunganRS"
        .TextMatrix(0, 9) = "JmlPembebasan"
        .TextMatrix(0, 10) = "KdPelayananRS"
        .TextMatrix(0, 11) = "KdRuangan"
        .TextMatrix(0, 12) = "KdInstalasi"
        .TextMatrix(0, 13) = "KdKelas"
        .TextMatrix(0, 14) = "KdJenisTarif"
        .TextMatrix(0, 15) = "JmlPelayanan"
        .TextMatrix(0, 16) = "StatusCITO"
        .TextMatrix(0, 17) = "IdPegawai"

        .ColWidth(0) = 1400
        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(1) = 2500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2200
        .ColWidth(4) = 1100
        .ColAlignment(4) = flexAlignCenterCenter
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0

    End With

End Sub

Private Sub cmdPerbaiki_Click()
    On Error GoTo hell_

    With fgData
        If fgData.TextMatrix(1, 0) = "" Then Exit Sub
        If MsgBox("Yakin akan memperbaiki nama pelayanan : " & .TextMatrix(.Row, 2) & " dengan tanggal pelayanan :" & .TextMatrix(.Row, 3), vbInformation + vbYesNo, "validasi") = vbNo Then Exit Sub
        If .TextMatrix(.Row, 12) = "09" Or .TextMatrix(.Row, 12) = "10" Or .TextMatrix(.Row, 12) = "16" Then
            sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(.Row, 11) & "'  and TglPelayanan ='" & Format(.TextMatrix(.Row, 3), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & .TextMatrix(.Row, 10) & "' " ' "
            Call msubRecFO(rsQuery, sqlQuery)

            If Add_TempHargaKomponenForPenunjang(.TextMatrix(.Row, 11), .TextMatrix(.Row, 3), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), .TextMatrix(.Row, 15), .TextMatrix(.Row, 16), "", "") = False Then Exit Sub

        ElseIf .TextMatrix(.Row, 11) = "401" Or .TextMatrix(.Row, 11) = "402" Or .TextMatrix(.Row, 11) = "403" Then
            Set rs = Nothing
            strSQL = "SELECT NoPendaftaran FROM  DokterPelaksanaOperasi" & _
            " where NoPendaftaran = '" & mstrNoPen & "' "
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(.Row, 11) & "'  and TglPelayanan ='" & Format(.TextMatrix(.Row, 3), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & .TextMatrix(.Row, 10) & "' " ' "
                Call msubRecFO(rsQuery, sqlQuery)
                If Add_TempHargaKomponenForIBS_DBNew(.TextMatrix(.Row, 11), .TextMatrix(.Row, 3), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), .TextMatrix(.Row, 14), .TextMatrix(.Row, 15), "") = False Then Exit Sub
            Else
                sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(.Row, 11) & "'  and TglPelayanan ='" & Format(.TextMatrix(.Row, 3), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & .TextMatrix(.Row, 10) & "' " ' "
                Call msubRecFO(rsQuery, sqlQuery)
                If Add_TempHargaKomponenForIBSNew(.TextMatrix(.Row, 11), .TextMatrix(.Row, 3), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), .TextMatrix(.Row, 14), .TextMatrix(.Row, 15), "") = False Then Exit Sub
            End If

        Else
            sqlQuery = "Delete TempHargaKomponen where NoPendaftaran ='" & mstrNoPen & "' AND KdRuangan ='" & .TextMatrix(.Row, 11) & "'  and TglPelayanan ='" & Format(.TextMatrix(.Row, 3), "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS = '" & .TextMatrix(.Row, 10) & "' " ' "
            Call msubRecFO(rsQuery, sqlQuery)
            If Add_TempHargaKomponen(.TextMatrix(.Row, 11), .TextMatrix(.Row, 3), .TextMatrix(.Row, 10), .TextMatrix(.Row, 13), .TextMatrix(.Row, 14), .TextMatrix(.Row, 6), .TextMatrix(.Row, 15), .TextMatrix(.Row, 16), .TextMatrix(.Row, 17), "") = False Then Exit Sub

        End If

        Call subLoadData
        cmdValidasiData.SetFocus

    End With
    Exit Sub
hell_:
    msubPesanError

End Sub

Private Sub cmdTutup_Click()
    Unload Me
    frmKeluarKamar.cmdSimpan.Enabled = True
End Sub

Private Sub cmdValidasiData_Click()
    On Error Resume Next
    Set rsQuery = Nothing
    If fgData.TextMatrix(1, 0) = "" Then Exit Sub
    pbData.Max = rs.RecordCount
    For i = 1 To pbData.Max
        DoEvents
        With fgData
            sqlQuery = "SELECT dbo.FB_TakeStatusDataValid('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 11) & "','" & .TextMatrix(i, 10) & "', '" & Format(.TextMatrix(i, 3), "yyyy/mm/dd HH:mm:ss ") & "', " & msubKonversiKomaTitik(.TextMatrix(i, 5)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 6)) & " , " & msubKonversiKomaTitik(.TextMatrix(i, 7)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 8)) & ", " & msubKonversiKomaTitik(.TextMatrix(i, 9)) & " )  as StatusData"
            Call msubRecFO(rsQuery, sqlQuery)
            strStatusData = rsQuery.Fields(0).Value
            .TextMatrix(i, 4) = strStatusData

            For j = 0 To 4
                .Col = j
                If .Col = 4 Then
                    If .TextMatrix(i, 4) = "T" Then
                        .Row = i
                        .CellBackColor = vbRed
                        .CellForeColor = vbWhite
                    End If
                End If
            Next j
        End With
        pbData.Value = Int(pbData.Value) + 1
    Next i

    'this is new from johnecholsphrasetyo@yahoo.co.id on 2009-04-29
    Set rsQuery = Nothing
    sqlQuery = "select NoPendaftaran from PasienDaftar where NoPendaftaran = '" & mstrNoPen & "' and ((kdKelompokPasien between '22' and '33') or (kdKelompokPasien between '39' and '46'))"
    Call msubRecFO(rsQuery, sqlQuery)
    If Not rsQuery.EOF Then
        If PacketRefreshingInsuranceOne(mstrNoPen) = False Then Exit Sub
        If PacketRefreshingInsuranceTwo(mstrNoPen) = False Then Exit Sub
    End If

End Sub

Private Function PacketRefreshingInsuranceOne(RegNo As String) As Boolean

    On Error GoTo StatusErr

    MousePointer = vbHourglass
    PacketRefreshingInsuranceOne = True
    Set dbcmd = New ADODB.Command
    With dbcmd

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(Len(Trim(RegNo)) = 0, Null, RegNo))

        .ActiveConnection = dbConn
        .CommandText = "Update_BiayaPelayananOnUbahJenisPasienNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validfasi"
            PacketRefreshingInsuranceOne = False
        Else
            Call Add_HistoryLoginActivity("Update_BiayaPelayananOnUbahJenisPasienNew")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault

    Exit Function

StatusErr:
    Set dbcmd = Nothing
    MousePointer = vbDefault
    PacketRefreshingInsuranceOne = False
    Call msubPesanError("PacketRefreshingInsuranceOne")
    MsgBox "Ulangi proses simpan", vbCritical, "Validasi"

End Function

Private Function PacketRefreshingInsuranceTwo(RegNo As String) As Boolean

    On Error GoTo StatusErr

    MousePointer = vbHourglass
    PacketRefreshingInsuranceTwo = True
    Set dbcmd = New ADODB.Command
    With dbcmd

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, IIf(Len(Trim(RegNo)) = 0, Null, RegNo))

        .ActiveConnection = dbConn
        .CommandText = "Add_DetailBiayaPelayananOnUbahJenisPasienNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validfasi"
            PacketRefreshingInsuranceTwo = False
        Else
            Call Add_HistoryLoginActivity("Add_DetailBiayaPelayananOnUbahJenisPasienNew")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    MousePointer = vbDefault

    Exit Function

StatusErr:
    Set dbcmd = Nothing
    MousePointer = vbDefault
    PacketRefreshingInsuranceTwo = False
    Call msubPesanError("PacketRefreshingInsuranceTwo")
    MsgBox "Ulangi proses simpan", vbCritical, "Validasi"

End Function

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)

    Select Case KeyCode
        Case 13
            If fgData.TextMatrix(fgData.Row, 2) = "" Then Exit Sub
            Call subLoadText
            txtIsi.Text = Trim(fgData.TextMatrix(fgData.Row, fgData.Col))
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadData
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then

        fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
        txtIsi.Visible = False

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        fgData.SetFocus
    End If

End Sub

Private Function Add_TempHargaKomponenForPenunjang(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_TarifCito As Integer, f_JmlPelayanan As Integer, f_StatusCito As String, f_Kdlaboratory As String, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Add_TempHargaKomponenForPenunjang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, f_StatusCito)
        .Parameters.Append .CreateParameter("KdLaboratory", adChar, adParamInput, 3, IIf(f_Kdlaboratory = "", Null, f_Kdlaboratory))
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(f_KdRuanganAsal = "", Null, f_KdRuanganAsal))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForPenunjangMNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            Add_TempHargaKomponenForPenunjang = False
        Else
            Call Add_HistoryLoginActivity("Add_TempHargaKomponenForPenunjangMNew")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    Add_TempHargaKomponenForPenunjang = False
    Call msubPesanError

End Function

Private Function Add_TempHargaKomponen(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_TarifCito As Integer, f_JmlPelayanan As Integer, f_StatusCito As String, f_kdDokter As String, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Add_TempHargaKomponen = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, f_StatusCito)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_kdDokter)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(f_KdRuanganAsal = "", Null, f_KdRuanganAsal))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            Add_TempHargaKomponen = False
        Else
            Call Add_HistoryLoginActivity("Add_TempHargaKomponenNew")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    Add_TempHargaKomponen = False
    Call msubPesanError

End Function

Private Function Add_TempHargaKomponenForIBS_DBNew(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_JmlPelayanan As Integer, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Add_TempHargaKomponenForIBS_DBNew = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(f_KdRuanganAsal = "", Null, f_KdRuanganAsal))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForIBS_DBNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            Add_TempHargaKomponenForIBS_DBNew = False
        Else
            Call Add_HistoryLoginActivity("Add_TempHargaKomponenForIBS_DBNew")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    Add_TempHargaKomponenForIBS_DBNew = False
    Call msubPesanError

End Function

Private Function Add_TempHargaKomponenForIBSNew(f_KdRuangan As String, f_tglPelayanan As Date, f_KdPelayananRS As String, f_KdKelas As String, f_KdJenisTarif As String, f_JmlPelayanan As Integer, f_KdRuanganAsal As String) As Boolean
    On Error GoTo errLoad

    Add_TempHargaKomponenForIBSNew = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, f_KdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, f_KdJenisTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , f_JmlPelayanan)
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, IIf(f_KdRuanganAsal = "", Null, f_KdRuanganAsal))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_TempHargaKomponenForIBSNew"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            Add_TempHargaKomponenForIBSNew = False
        Else
            Call Add_HistoryLoginActivity("Add_TempHargaKomponenForIBSNew")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)

    Exit Function
errLoad:
    Add_TempHargaKomponenForIBSNew = False
    Call msubPesanError

End Function

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

