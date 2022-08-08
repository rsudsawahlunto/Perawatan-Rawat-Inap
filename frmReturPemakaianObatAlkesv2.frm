VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReturPemakaianObatAlkesv2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Retur Pemakaian Obat dan Alat Kesehatan"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReturPemakaianObatAlkesv2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13905
   Begin VB.TextBox txtCariData 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4320
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4095
   End
   Begin VB.TextBox txtNoRetur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      MaxLength       =   15
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtKeterangan 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6360
      Width           =   4095
   End
   Begin VB.TextBox txtIsi 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTotalRetur 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      TabIndex        =   11
      Top             =   6840
      Width           =   2775
   End
   Begin VB.TextBox txtTotalBiaya 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      TabIndex        =   10
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   10320
      TabIndex        =   8
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   12120
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Resep"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   -720
      Visible         =   0   'False
      Width           =   13815
      Begin VB.CheckBox chkNoResep 
         Caption         =   "No Resep"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNoResep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtDokter 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   3600
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dtpTglResep 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   106627075
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcTglPelayanan 
         Height          =   330
         Left            =   11160
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dokter Penulis Resep"
         Height          =   210
         Index           =   2
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Pelayanan"
         Height          =   210
         Index           =   0
         Left            =   11160
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Resep"
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   870
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgData 
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpTglRetur 
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   117571587
      UpDown          =   -1  'True
      CurrentDate     =   37760
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   23
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
      Picture         =   "frmReturPemakaianObatAlkesv2.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmReturPemakaianObatAlkesv2.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmReturPemakaianObatAlkesv2.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pencarian Data (No Resep)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Width           =   4125
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13800
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl. Retur"
      Height          =   210
      Index           =   6
      Left            =   135
      TabIndex        =   19
      Top             =   6450
      Width           =   825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   18
      Top             =   6360
      Width           =   945
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Retur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9480
      TabIndex        =   17
      Top             =   6840
      Width           =   1410
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Biaya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   9480
      TabIndex        =   16
      Top             =   6360
      Width           =   1395
   End
End
Attribute VB_Name = "frmReturPemakaianObatAlkesv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim substrNomorRetur As String
Dim subbolSimpan As Boolean
Dim i As Integer

Public Sub subLoadReturPemakaian(Optional s_Kriteria As String)
    On Error GoTo errLoad
    Dim i As Integer
    Dim j As Integer
    Dim intTotalBiaya As Currency
    Dim tempTotalBiaya As String
    fgData.Cols = 17
    txtTotalBiaya.Text = 0
    txtTotalRetur.Text = 0
    intTotalBiaya = 0

    Call subSetGrid
    strSQL = "SELECT NoRacikan,NoResep, TglResep, DokterPenulisResep, JenisBarang, NamaBarang, AsalBarang, HargaSatuan, JmlBarang, KdRuangan, KdBarang, KdAsal, SatuanJml, TglPelayanan, JmlService, TarifService, HargaSebelumTarifService, NoTerima, ResepKe,KdJenisObat" & _
    " FROM V_RiwayatPemakaianObatAlkes" & _
    " WHERE NoPendaftaran = '" & mstrNoPen & "' AND KdRuangan = '" & mstrKdRuangan & "'" & s_Kriteria & " " & _
    " ORDER BY NamaBarang, NoTerima,TglPelayanan Desc"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = True Then
    
        
        With fgData

            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
            .TextMatrix(1, 8) = ""
            .TextMatrix(1, 9) = ""
            .TextMatrix(1, 10) = ""
            .TextMatrix(1, 11) = ""
            .TextMatrix(1, 12) = ""
            .TextMatrix(1, 13) = ""
            .TextMatrix(1, 14) = ""
            .TextMatrix(1, 15) = ""
            .TextMatrix(1, 16) = ""
        End With

    txtTotalBiaya.Text = "0"
    
    
    Else

    mstrKdRuanganPasien = rs("KdRuangan")
    txtNoResep.Text = IIf(IsNull(rs("NoResep")), "", rs("NoResep"))
    dtpTglResep.value = IIf(IsNull(rs("TglResep")), Now, rs("TglResep"))
    txtDokter.Text = IIf(IsNull(rs("DokterPenulisResep")), "", rs("DokterPenulisResep"))
    j = 1
    For i = 1 To rs.RecordCount
        With fgData
            If (rs("JmlBarang") <> 0) Then
            .TextMatrix(j, 0) = rs("TglPelayanan")
            .TextMatrix(j, 1) = IIf(IsNull(rs("NoResep")), "", rs("NoResep"))
            .TextMatrix(j, 2) = rs("NamaBarang")
            .TextMatrix(j, 3) = rs("AsalBarang")
            .TextMatrix(j, 4) = IIf(rs("HargaSatuan") = 0, 0, FormatPembulatan(CDbl(rs("HargaSatuan")), mstrKdInstalasiLogin))
            .TextMatrix(j, 5) = rs("JmlBarang")
            .TextMatrix(j, 6) = 0
            .TextMatrix(j, 7) = 0
            .TextMatrix(j, 8) = rs("KdBarang")
            .TextMatrix(j, 9) = rs("KdAsal")
            .TextMatrix(j, 10) = rs("SatuanJml")
            .TextMatrix(j, 11) = rs("JmlService")
            .TextMatrix(j, 12) = rs("TarifService")
            .TextMatrix(j, 13) = rs("HargaSebelumTarifService")
            .TextMatrix(j, 14) = rs("NoTerima")
            .TextMatrix(j, 15) = rs("ResepKe")
            .TextMatrix(j, 16) = rs("KdJenisObat")
            .TextMatrix(j, 17) = IIf(IsNull(rs("NoResep")), "", rs("NoResep"))  'rs("NoResep")
            .TextMatrix(j, 18) = IIf(IsNull(rs("NoRacikan")), "", rs("NoRacikan"))

            intTotalBiaya = intTotalBiaya + ((CCur(.TextMatrix(j, 5)) * CCur(.TextMatrix(j, 13))) + (CCur(.TextMatrix(j, 11)) * CCur(.TextMatrix(j, 12))))

            .Rows = .Rows + 1
            j = j + 1
            End If
            rs.MoveNext
        End With
    Next i

    For j = 1 To Len(intTotalBiaya)
        tempTotalBiaya = Mid(intTotalBiaya, j, 1)
        If tempTotalBiaya = "," Then tempTotalBiaya = "."
        txtTotalBiaya.Text = txtTotalBiaya.Text & tempTotalBiaya
    Next j

    txtTotalBiaya.Text = FormatPembulatan(CDbl(txtTotalBiaya.Text), mstrKdInstalasiLogin)
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcTglPelayanan, rs, "SELECT DISTINCT TglPelayanan, TglPelayanan AS Alias FROM V_RiwayatPemakaianObatAlkes WHERE NoPendaftaran = '" & mstrNoPen & "' ORDER BY TglPelayanan")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    txtIsi.Left = fgData.Left
    Select Case fgData.Col
        Case 6
            txtIsi.MaxLength = 500
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
        .Cols = 19

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "Tanggal"
        .TextMatrix(0, 1) = "No Resep"
        .TextMatrix(0, 2) = "Nama Barang"
        .TextMatrix(0, 3) = "Asal Barang"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Jumlah"
        .TextMatrix(0, 6) = "Jumlah Retur"
        .TextMatrix(0, 7) = "Total"
        .TextMatrix(0, 8) = "KdBarang"
        .TextMatrix(0, 9) = "KdAsal"
        .TextMatrix(0, 10) = "Satuan"
        .TextMatrix(0, 11) = "JmlService"
        .TextMatrix(0, 12) = "TarifService"
        .TextMatrix(0, 13) = "HargaSebelumTarifService"
        .TextMatrix(0, 14) = "NoTerima"
        .TextMatrix(0, 15) = "Rke"
        .TextMatrix(0, 16) = "JenisObat"
        .TextMatrix(0, 17) = "NoResep"
        .TextMatrix(0, 18) = "NoRacikan"

        .ColWidth(0) = 2000
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 1500
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 1000
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0

    End With
End Sub

Private Sub chkNoResep_Click()
    If chkNoResep.value = vbChecked Then
        txtNoResep.Enabled = True
    Else
        txtNoResep.Enabled = False
    End If
End Sub

Private Sub chkNoResep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If txtNoResep.Enabled = True Then txtNoResep.SetFocus Else fgData.SetFocus
End Sub

Private Sub cmdSimpan_Click()
    Dim i As Integer
    
    If chkNoResep.value = vbChecked Then If Periksa("text", txtNoResep, "No Resep kosong") = False Then Exit Sub
    For i = 1 To fgData.Rows - 1
     ' If fgData.TextMatrix(i, 1) = "" Then MsgBox "Barang yang diretur harus ada"
     ' Exit Sub
    Next i
    
    
   ' Call txtIsi_KeyPress(13)
    If Val(txtTotalRetur) = 0 Then
        MsgBox "Minimal 1 barang yang diretur", vbExclamation, "Validasi"
        fgData.SetFocus: fgData.Col = 5
        Exit Sub
    End If

    If sp_Retur() = False Then Exit Sub
    If substrNomorRetur = "" Then Exit Sub
    For i = 1 To fgData.Rows - 1
        With fgData
            If Val(.TextMatrix(i, 6)) <> 0 Then If sp_ReturPemakaianObatAlkes(.TextMatrix(i, 8), .TextMatrix(i, 9), .TextMatrix(i, 0), .TextMatrix(i, 10), .TextMatrix(i, 6), .TextMatrix(i, 14), .TextMatrix(i, 15)) = False Then Exit Sub
        End With
    Next i

    Call Add_HistoryLoginActivity("Add_Retur+Add_ReturnPemakaianObatAlkes")
    subbolSimpan = True
    
    Call subLoadReturPemakaian
    MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
    cmdSimpan.Enabled = False
    cmdTutup.SetFocus
End Sub

Private Sub cmdTutup_Click()
    If subbolSimpan = False Then
        If MsgBox("Simpan data retur pemakaian obat dan alat kesehatan", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
    Call frmTransaksiPasien.subPemakaianObatAlkes
'    Call frmTransaksiPasien.subLoadReturPemakaianObatAlkes
End Sub

Private Sub dcTglPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkNoResep.SetFocus
End Sub

Private Sub dtpTglResep_Change()
    dtpTglResep.MaxDate = Now
End Sub

Private Sub dtpTglRetur_Change()
    dtpTglRetur.MaxDate = Now
End Sub

Private Sub dtpTglRetur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

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

        Case vbKeyDelete
            If fgData.Row = fgData.Rows - 1 Then Exit Sub
            fgData.RemoveItem fgData.Row

    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    dtpTglResep.value = Now
    dtpTglRetur.value = Now
    Call subSetGrid
    Call subLoadDcSource

    Call subLoadReturPemakaian
    subbolSimpan = False
End Sub

Private Sub txtCariData_Change()
    On Error GoTo errLoad

    Call subLoadReturPemakaian(" AND NoResep LIKE '%" & txtCariData.Text & "%'")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i, intRow As Integer
    Dim dblJmlBrg, dblSelisih As Double
    Dim strKdBrg, strKdAsal, strSatJml As String
    Dim dtTglPel As Date

    If KeyAscii = 13 Then
        If Val(txtIsi.Text) = 0 Then
            txtIsi.Text = 0
            fgData.CellForeColor = vbBlack: fgData.CellFontBold = False
        Else
            fgData.CellForeColor = vbBlue: fgData.CellFontBold = True
        End If

        If Val(txtIsi.Text) > Val(fgData.TextMatrix(fgData.Row, 5)) Then
            MsgBox "Jumlah retur lebih besar dari jumlah asal", vbExclamation, "Validasi"
            txtIsi.SelStart = 0
            txtIsi.SelLength = Len(txtIsi.Text)
            Exit Sub
        End If
        strSQL = "select value from settingglobal where Prefix='KdJenisObatRacikan'"
        Call msubRecFO(rsK, strSQL)
        If (rsK.EOF = False) Then
            If (rsK(0).value = Val(fgData.TextMatrix(fgData.Row, 16))) Then
                fgData.TextMatrix(fgData.Row, fgData.Col) = fgData.TextMatrix(fgData.Row, 5)
                Dim j As Integer
                For j = 1 To fgData.Rows - 2
                    If (fgData.TextMatrix(j, 17) = fgData.TextMatrix(fgData.Row, 17) And fgData.TextMatrix(j, 18) = fgData.TextMatrix(fgData.Row, 18)) Then
                        If (txtIsi.Text = "0") Then
                            fgData.TextMatrix(j, fgData.Col) = "0"
                        Else
                            fgData.TextMatrix(j, fgData.Col) = fgData.TextMatrix(j, 5)
                        End If
                        
                        fgData.TextMatrix(j, 7) = (fgData.TextMatrix(j, fgData.Col) * CCur(fgData.TextMatrix(j, 13))) + (CCur(fgData.TextMatrix(j, 11)) * CCur(fgData.TextMatrix(j, 12)))
                        If Val(fgData.TextMatrix(j, 7)) >= 0 Then fgData.TextMatrix(j, 7) = FormatPembulatan(CDbl(fgData.TextMatrix(j, 7)), mstrKdInstalasiLogin)
                    End If
                    
                    
                Next j
            Else
                fgData.TextMatrix(fgData.Row, fgData.Col) = msubKonversiKomaTitik(txtIsi.Text)
            End If
        Else
            fgData.TextMatrix(fgData.Row, fgData.Col) = msubKonversiKomaTitik(txtIsi.Text)
        End If
        'if(fgData.TextMatrix(fgData.Row, 16))
        'konvert koma col jumlah
        

        fgData.TextMatrix(fgData.Row, 7) = (fgData.TextMatrix(fgData.Row, fgData.Col) * CCur(fgData.TextMatrix(fgData.Row, 13))) + (CCur(fgData.TextMatrix(fgData.Row, 11)) * CCur(fgData.TextMatrix(fgData.Row, 12)))
        If Val(fgData.TextMatrix(fgData.Row, 7)) > 0 Then fgData.TextMatrix(fgData.Row, 7) = FormatPembulatan(CDbl(fgData.TextMatrix(fgData.Row, 7)), mstrKdInstalasiLogin)
        txtIsi.Visible = False

        txtTotalRetur.Text = FormatPembulatan(CDbl(HitungTotalRetur), mstrKdInstalasiLogin)

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
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Function sp_Retur() As Boolean
    On Error GoTo errLoad

    sp_Retur = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, txtNoRetur.Text)
        .Parameters.Append .CreateParameter("TglRetur", adDate, adParamInput, , Format(dtpTglRetur.value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 50, IIf(Len(Trim(txtKeterangan.Text)) = 0, Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("OutputNoRetur", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_Retur"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Retur", vbCritical, "Validasi"
            sp_Retur = False
        Else
            txtNoRetur.Text = .Parameters("OutputNoRetur").value
            substrNomorRetur = .Parameters("OutputNoRetur").value

        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    sp_Retur = False
    Call msubPesanError
End Function

Private Function sp_ReturPemakaianObatAlkes(f_KdBarang As String, f_KdAsal As String, f_TglPelayanan As Date, f_Satuan As String, f_JumlahRetur As Double, f_NoTerima, f_ResepKe As Integer) As Boolean
    On Error GoTo errLoad

    sp_ReturPemakaianObatAlkes = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRetur", adChar, adParamInput, 10, substrNomorRetur)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuanganPasien)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_TglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, f_Satuan)
        .Parameters.Append .CreateParameter("JmlRetur", adDouble, adParamInput, , CDbl(f_JumlahRetur))
        .Parameters.Append .CreateParameter("NoTerima", adChar, adParamInput, 10, f_NoTerima)
        .Parameters.Append .CreateParameter("ResepKe", adTinyInt, adParamInput, , f_ResepKe)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_ReturnPemakaianObatAlkes"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data Retur Pemakaian Obat dan Alat Kesehatan", vbCritical, "Validasi"
            sp_ReturPemakaianObatAlkes = False

        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    Call msubPesanError
End Function

Private Function HitungTotalRetur() As Currency
    Dim i As Integer

    HitungTotalRetur = 0
    For i = 1 To fgData.Rows - 2
        HitungTotalRetur = HitungTotalRetur + fgData.TextMatrix(i, 7)
    Next i
End Function

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
