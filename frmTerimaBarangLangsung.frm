VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmTerimaBarangLangsung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Terima Barang Langsung"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTerimaBarangLangsung.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12420
   Begin MSDataGridLib.DataGrid dgObatAlkes 
      Height          =   2535
      Left            =   1200
      TabIndex        =   7
      Top             =   2800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTotalDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7200
      MaxLength       =   12
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   7560
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dgNamaPenerima 
      Height          =   2535
      Left            =   12720
      TabIndex        =   22
      Top             =   5760
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtTotalBiaya 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10320
      MaxLength       =   12
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Terima Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   1200
      Width           =   12375
      Begin VB.TextBox txtKdUserPenerima 
         Height          =   315
         Left            =   10440
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtNamaPenerima 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   10800
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtNoKirim 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpTglKirim 
         Height          =   330
         Left            =   2040
         TabIndex        =   1
         Top             =   480
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
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   16646147
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcStatusBarang 
         Height          =   330
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dcRuanganPengirim 
         Height          =   330
         Left            =   6120
         TabIndex        =   26
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruangan Pengirim"
         Height          =   210
         Index           =   4
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penerima"
         Height          =   210
         Index           =   2
         Left            =   10800
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   210
         Index           =   11
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Terima"
         Height          =   210
         Index           =   8
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Terima"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame Frame0 
      Caption         =   "Data Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   12375
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txtKdSatuan 
         Height          =   315
         Left            =   3840
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtKdBarang 
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4575
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8070
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   7440
      Width           =   1695
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
      Left            =   10560
      Picture         =   "frmTerimaBarangLangsung.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTerimaBarangLangsung.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmTerimaBarangLangsung.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   5880
      TabIndex        =   25
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Biaya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   9240
      TabIndex        =   19
      Top             =   7560
      Width           =   945
   End
End
Attribute VB_Name = "frmTerimaBarangLangsung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim substrNoOrder As String
'Dim substrKdPegawai As String
'
'
''Private Sub chkNoOrder_Click()
''    If chkNoOrder.Value = vbChecked Then
''        txtNoOrder.Enabled = True
''        txtNoOrder.Text = Format(Now, "yy") & Format(Now, "MM") & Format(Now, "dd")
''        txtNoOrder.SelStart = Len(txtNoOrder.Text)
''    Else
''        txtNoOrder.Enabled = False
''
''        dtpTglOrder.Value = Now
''        txtRuanganPemesan.Text = ""
''        txtNamaPemesan.Text = ""
''        dcRuanganPenerima.BoundText = ""
''    End If
''End Sub
'
''Private Sub chkNoOrder_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        If chkNoOrder.Value = vbChecked Then txtNoOrder.SetFocus Else dtpTglKirim.SetFocus
''    End If
''End Sub
'
'Private Sub cmdBatal_Click()
'    Call subKosong
'    Call subLoadDcSource
'    Call subSetGrid
'End Sub
'
'Private Sub cmdSimpan_Click()
'On Error GoTo errLoad
'Dim i As Integer
'
'    If Periksa("datacombo", dcRuanganPengirim, "Nama ruangan pengirim kosong") = False Then Exit Sub
'
'    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data barang harus diisi", vbExclamation, "Validasi": Exit Sub
'
'    For i = 1 To fgData.Rows - 2
'        With fgData
'            If .TextMatrix(i, 5) = 0 Or .TextMatrix(i, 5) = "" Then
'                MsgBox "Qty barang tidak boleh nol", vbExclamation, "Validasi"
'                .SetFocus: .Row = i: .Col = 5
'                Exit Sub
'            End If
'        End With
'    Next i
'
'    If sp_StrukKirim() = False Then Exit Sub
'    For i = 1 To fgData.Rows - 1
'        With fgData
'            If .TextMatrix(i, 0) <> "" Then
'                If sp_DetailBarangKeluar(.TextMatrix(i, 0), .TextMatrix(i, 9), .TextMatrix(i, 5), .TextMatrix(i, 6), _
'                    .TextMatrix(i, 7), 0, "A") = False Then Exit Sub
'            End If
'        End With
'    Next i
'
'    Call Add_HistoryLoginActivity("Add_StrukKirim+Add_DetailBarangKeluar")
'    MsgBox "No Terima : " & txtNoKirim.Text, vbInformation, "Informasi"
'    Call cmdBatal_Click
''    subbolSimpan = True
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub cmdTutup_Click()
'    'If subbolSimpan = False Then
''        If MsgBox("Simpan Data Penerimaan Barang?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
''            Call cmdSimpan_Click
''            Exit Sub
''        End If
'    'End If
' '   If txtNamaFormPengirim.Text = "frmDaftarPemesananBarangdariRuangan" Then frmDaftarPemesananBarangdariRuangan.cmdCari_Click
'    Unload Me
'End Sub
'
'Private Sub Hapus()
'On Error GoTo errLoad
'Dim i As Integer
'    With fgData
'        If .Row = .Rows Then Exit Sub
'        If .Row = 0 Then Exit Sub
'
'        If .Rows = 2 Then
'            For i = 0 To .Cols - 1
'                .TextMatrix(1, i) = ""
'            Next i
'            Exit Sub
'        Else
'            .RemoveItem .Row
'        End If
'    End With
'    Call subHitungTotal
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dcAsalBarang_Change()
'On Error GoTo errLoad
'Dim j As Integer
'Dim tempDiscount As String
'
'    If fgData.TextMatrix(fgData.Row, 0) = "" Then Exit Sub
'        strSQL = "SELECT Satuan,JmlStok, HargaNetto, Discount" & _
'            " From V_CariBarangMedis" & _
'            " WHERE (KdBarang = '" & txtKdBarang.Text & "') AND (KdAsal = '" & dcAsalBarang.BoundText & "') AND (KdRuangan = '" & mstrKdRuangan & "')"
'
'    Call msubRecFO(rs, strSQL)
'
'    With fgData
'        If rs.EOF = True Then
'            .TextMatrix(.Row, 4) = 0    'JmlStok
'            .TextMatrix(.Row, 5) = 0    'JmlKirim
'            .TextMatrix(.Row, 6) = 0    'HargaSatuan
'            .TextMatrix(.Row, 7) = 0    'Discount
'        Else
'            .TextMatrix(.Row, 3) = rs("Satuan").Value
'            .TextMatrix(.Row, 4) = rs("JmlStok").Value
'            .TextMatrix(.Row, 5) = 0
'            .TextMatrix(.Row, 6) = IIf(IsNull(rs("HargaNetto")), 0, rs("HargaNetto"))
'            .TextMatrix(.Row, 7) = 0
'
'
'            If Not IsNull(rs("Discount")) Then
'                For j = 1 To Len(rs("Discount"))
'                    tempDiscount = Mid(rs("Discount").Value, j, 1)
'                    If tempDiscount = "," Then tempDiscount = "."
'                    .TextMatrix(.Row, 8) = .TextMatrix(.Row, 8) & tempDiscount
'                Next j
'            End If
'        End If
'    End With
'
'    fgData.TextMatrix(fgData.Row, 2) = dcAsalBarang.Text
'    fgData.TextMatrix(fgData.Row, 9) = dcAsalBarang.BoundText
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dcAsalBarang_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then dcAsalBarang.Visible = False: fgData.SetFocus
'End Sub
'
'Private Sub dcAsalBarang_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Or KeyAscii = 27 Then
'        Call dcAsalBarang_Change
'        dcAsalBarang.Visible = False
'        fgData.Col = 5
'        fgData.SetFocus
'    End If
'End Sub
'
'Private Sub dcAsalBarang_LostFocus()
'    dcAsalBarang.Visible = False
'End Sub
'
''Private Sub dcRuanganPenerima_KeyPress(KeyAscii As Integer)
''On Error GoTo errLoad
''
''    If KeyAscii = 13 Then
''        If Len(Trim(dcRuanganPenerima.Text)) = 0 Then fgData.SetFocus: fgData.Col = 0: Exit Sub
''        If dcRuanganPenerima.MatchedWithList = True Then fgData.SetFocus: fgData.Col = 0: Exit Sub
''        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcRuanganPenerima.Text & "%'")
''        If dbRst.EOF = True Then Exit Sub
''        dcRuanganPenerima.BoundText = dbRst(0).Value
''        dcRuanganPenerima.Text = dbRst(1).Value
''    End If
''Exit Sub
''errLoad:
''    Call msubPesanError
''End Sub
'
'Private Sub dcRuanganPengirim_KeyPress(KeyAscii As Integer)
'On Error GoTo errLoad
'
'    If KeyAscii = 13 Then
'        If Len(Trim(dcRuanganPengirim.Text)) = 0 Then fgData.SetFocus: fgData.Col = 1: Exit Sub
'        If dcRuanganPengirim.MatchedWithList = True Then fgData.SetFocus: fgData.Col = 1: Exit Sub
'        Call msubRecFO(dbRst, "SELECT KdRuangan, NamaRuangan FROM Ruangan WHERE NamaRuangan LIKE '%" & dcRuanganPengirim.Text & "%'")
'        If dbRst.EOF = True Then Exit Sub
'        dcRuanganPengirim.BoundText = dbRst(0).Value
'        dcRuanganPengirim.Text = dbRst(1).Value
'    End If
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dgNamaPenerima_DblClick()
'On Error GoTo errLoad
'    If dgNamaPenerima.ApproxCount = 0 Then Exit Sub
'    txtKdUserPenerima.Text = dgNamaPenerima.Columns("IdPegawai").Value
'    txtNamaPenerima.Text = dgNamaPenerima.Columns("Nama Pemeriksa").Value
'    substrKdPegawai = dgNamaPenerima.Columns("IdPegawai").Value
'    dgNamaPenerima.Visible = False
'    fgData.SetFocus
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub dgNamaPenerima_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Call dgNamaPenerima_DblClick
'End Sub
'
'Private Sub dgObatAlkes_DblClick()
'On Error GoTo errLoad
''Dim j As Integer
''Dim tempDiscount As String
'    With fgData
'        .TextMatrix(.Row, 0) = dgObatAlkes.Columns("KdBarang")
'        txtKdBarang = dgObatAlkes.Columns("KdBarang")
'        txtKdSatuan.Text = dgObatAlkes.Columns("KdSatuanJmlB")
'        .TextMatrix(.Row, 1) = dgObatAlkes.Columns("Nama Barang")
'        .TextMatrix(.Row, 2) = dgObatAlkes.Columns("Asal Barang")
'        .TextMatrix(.Row, 3) = dgObatAlkes.Columns("Satuan")
'        .TextMatrix(.Row, 4) = dgObatAlkes.Columns("JmlStok")
'        .TextMatrix(.Row, 5) = 0
'        .TextMatrix(.Row, 6) = IIf(dgObatAlkes.Columns("HargaNetto") = "", 0, dgObatAlkes.Columns("HargaNetto"))
'        .TextMatrix(.Row, 7) = 0
'        .TextMatrix(.Row, 9) = dgObatAlkes.Columns("KdAsal")
'
'        dgObatAlkes.Visible = False
'        .Col = 5
'        .SetFocus
'       ' dcAsalBarang.BoundText = .TextMatrix(.Row, 2)
'        'Call subLoadDataCombo(dcAsalBarang)
'    End With
'Exit Sub
'errLoad:
'End Sub
'
'Private Sub dgObatAlkes_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then dgObatAlkes.Visible = False: fgData.SetFocus
'End Sub
'
'Private Sub dgObatAlkes_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Call dgObatAlkes_DblClick
'End Sub
'
'Private Sub dtpTglKirim_Change()
'    dtpTglKirim.MaxDate = Now
'End Sub
'
'Private Sub dtpTglOrder_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dcRuanganPengirim.SetFocus
'End Sub
'
'Private Sub dtpTglKirim_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then dcRuanganPengirim.SetFocus
'End Sub
'
'Private Sub fgData_DblClick()
'    Call fgData_KeyDown(13, 0)
'End Sub
'
'Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim i As Integer
'    Select Case KeyCode
'        Case 13
'            If fgData.Col = fgData.Cols - 1 Then
'                If fgData.TextMatrix(fgData.Row, 2) <> "" Then
'                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then fgData.Rows = fgData.Rows + 1
'                    fgData.Row = fgData.Rows - 1
'                    fgData.Col = 1
'                Else
'                    fgData.Col = 1
'                End If
'            Else
'                For i = 0 To fgData.Cols - 2
'                    If fgData.Col = fgData.Cols - 1 Then Exit For
'                    fgData.Col = fgData.Col + 1
'                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
'                Next i
'            End If
'            fgData.SetFocus
'
'        Case 27
'            dgObatAlkes.Visible = False
'
'        Case vbKeyDelete
'            With fgData
'                If .Row = .Rows Then Exit Sub
'                If .Row = 0 Then Exit Sub
'
'                If .Rows = 2 Then
'                    For i = 0 To .Cols - 1
'                        .TextMatrix(1, i) = ""
'                    Next i
'                    Exit Sub
'                Else
'                    .RemoveItem .Row
'                End If
'            End With
'            Call subHitungTotal
'
'    End Select
'End Sub
'
'Private Sub fgData_KeyPress(KeyAscii As Integer)
'On Error GoTo errLoad
'
'    txtIsi.Text = ""
'    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
'        KeyAscii = 0
'        Exit Sub
'    End If
'
'    Select Case fgData.Col
''        Case 0 'kode barang
''            txtIsi.MaxLength = 9
''            Call subLoadText
''            txtIsi.Text = Chr(KeyAscii)
''            txtIsi.SelStart = Len(txtIsi.Text)
'
'        Case 1 'nama barang
'            txtIsi.MaxLength = 20
'            Call subLoadText
'            txtIsi.Text = Chr(KeyAscii)
'            txtIsi.SelStart = Len(txtIsi.Text)
'
'        Case 5 ' jml kirim
'            txtIsi.MaxLength = 4
'            Call subLoadText
'            txtIsi.Text = Chr(KeyAscii)
'            txtIsi.SelStart = Len(txtIsi.Text)
'
'        Case 2 'Asal Barang
'            fgData.Col = 2
'            Call subLoadDataCombo(dcAsalBarang)
'    End Select
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 39 Then KeyAscii = 0
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo errLoad
'
'    Call PlayFlashMovie(Me)
'    Call centerForm(Me, MDIUtama)
'
'    Call subKosong
'    Call subSetGrid
'    Call subLoadDcSource
'
'    dgObatAlkes.Top = 3720
'    dgObatAlkes.Left = 120
'    dgObatAlkes.Visible = False
'
'    dgNamaPenerima.Top = 2760
'    dgNamaPenerima.Left = 7800
'    dgNamaPenerima.Visible = False
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub txtIsi_Change()
'On Error GoTo errLoad
'Dim i As Integer
''    If txtIsi.Text = "" Then Exit Sub
'    Select Case fgData.Col
''        Case 0 'kode barang
''                strSQL = "select DISTINCT TOP 100 KdBarang, [Jenis Barang], [Nama Barang], [Asal Barang], Satuan, JmlStok, HargaNetto, Discount, KdSatuanJmlB, Kekuatan, KdAsal " & _
''                    " from V_AmbilStockBarang " & _
''                    " where KdBarang like '" & txtIsi & "%' AND KdRuangan = '" & mstrKdRuangan & "' ORDER BY KdBarang"
''
''            Call msubRecFO(dbRst, strSQL)
''            Set dgObatAlkes.DataSource = dbRst
''            With dgObatAlkes
''                .Columns("KdBarang").Width = 1250
''                .Columns("Jenis Barang").Width = 0 '1250
''                .Columns("Nama Barang").Width = 3900
''                .Columns("Satuan").Width = 0
''                .Columns("Kekuatan").Width = 1150
''                .Columns("KdSatuanJmlB").Width = 0
''                .Columns("KdAsal").Width = 0
''                .Left = fgData.Left
''                .Top = 3600
''                .Visible = True
''                For i = 1 To fgData.Row - 1
''                    .Top = .Top + fgData.RowHeight(i)
''                Next i
''                If fgData.TopRow > 1 Then
''                    .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
''                End If
''                .Top = .Top + 200
''            End With
'
'        Case 1 'nama barang
'                strSQL = "select [Nama Barang], [Asal Barang], [Jenis Barang], Satuan, JmlStok, KdBarang, KdSatuanJmlB, Kekuatan, KdAsal, HargaNetto" & _
'                    " from V_AmbilStockBarang " & _
'                    " where [Nama Barang] like '" & txtIsi & "%' AND KdRuangan = '" & mstrKdRuangan & "' ORDER BY [Nama Barang]"
'
'            Call msubRecFO(dbRst, strSQL)
'            Set dgObatAlkes.DataSource = dbRst
'            With dgObatAlkes
'                .Columns("Jenis Barang").Width = 0 '1250
'                .Columns("Nama Barang").Width = 3900
'                .Columns("Satuan").Width = 0
'                .Columns("Kekuatan").Width = 1150
'                .Columns("KdBarang").Width = 1250
'                .Columns("KdSatuanJmlB").Width = 0
'                .Columns("KdAsal").Width = 0
'                .Left = 1300
'                .Top = 3000
'                .Visible = True
'                For i = 1 To fgData.Row - 1
'                    .Top = .Top + fgData.RowHeight(i)
'                Next i
'                If fgData.TopRow > 1 Then
'                    .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
'                End If
'                .Top = .Top + 200
'            End With
'    End Select
'
'errLoad:
'End Sub
'
'Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDown Then If dgObatAlkes.Visible = True Then dgObatAlkes.SetFocus
'End Sub
'
'Private Sub txtIsi_KeyPress(KeyAscii As Integer)
'On Error GoTo errLoad
'Dim i As Integer
'    If KeyAscii = 13 Then
'        With fgData
'            Select Case .Col
''                Case 0
''                    If dgObatAlkes.Visible = True Then
''                        dgObatAlkes.SetFocus
''                        Exit Sub
''                    Else
''                        fgData.SetFocus
''                        fgData.Col = 1
''                    End If
'                Case 1
'                    If dgObatAlkes.Visible = True Then
'                        dgObatAlkes.SetFocus
'                        Exit Sub
'                    Else
'                        fgData.SetFocus
'                        fgData.Col = 2
'                    End If
'                Case 5
'                    If Val(txtIsi.Text) = 0 Then txtIsi.Text = 0
''                    If Val(txtIsi.Text) > Val(.TextMatrix(.Row, 3)) Then
''                        MsgBox "Jumlah lebih besar dari stock (" & .TextMatrix(.Row, 3) & ")", vbExclamation, "Validasi"
''                        txtIsi.SelStart = 0: txtIsi.SelLength = Len(txtIsi.Text)
''                        Exit Sub
''                    End If
'                    .TextMatrix(.Row, .Col) = msubKonversiKomaTitik(txtIsi.Text)
'                    .TextMatrix(.Row, 8) = CDec(.TextMatrix(.Row, 5)) * CDbl(.TextMatrix(.Row, 6))
'                    .TextMatrix(.Row, 10) = (Val(.TextMatrix(.Row, 7)) / 100) * (CDbl(.TextMatrix(.Row, 5)) * CDbl(.TextMatrix(.Row, 6)))
'
'                    Call subHitungTotal
'
'                    If .RowPos(.Row) >= .Height - 360 Then
'                        .SetFocus
'                        SendKeys "{DOWN}"
'                        Exit Sub
'                    End If
'                    .SetFocus
'                    If fgData.TextMatrix(fgData.Rows - 1, 2) <> "" Then fgData.Rows = fgData.Rows + 1
'                    fgData.SetFocus
'                    If txtNamaFormPengirim.Text <> "frmDaftarPengirimananAntarRuangan" Then
'                        fgData.Row = fgData.Rows - 1
'                        fgData.Col = 0
'                    Else
'                        fgData.Col = 5
'                    End If
'            End Select
'        End With
'    ElseIf KeyAscii = 27 Then
'        dgObatAlkes.Visible = False
'        txtIsi.Visible = False
'        fgData.SetFocus
'    End If
'
'    If fgData.Col = 5 Then
'        If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".")) Then KeyAscii = 0
'    End If
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub txtIsi_LostFocus()
'    txtIsi.Visible = False
'End Sub
'
'Private Sub subKosong()
'    txtNoKirim.Text = ""
'    dtpTglKirim.Value = Now
'    dcRuanganPengirim.BoundText = ""
'    dcStatusBarang.BoundText = ""
'    txtNamaPenerima.Text = ""
'    txtKdUserPenerima.Text = ""
'
'    substrNoOrder = ""
'    txtTotalBiaya.Text = 0
'    dgObatAlkes.Visible = False
'    dgNamaPenerima.Visible = False
'End Sub
'
'Private Sub subSetGrid()
'On Error GoTo errLoad
'    With fgData
'        .clear
'        .Rows = 2
'        .Cols = 11
'
'        .RowHeight(0) = 400
'
'        .TextMatrix(0, 0) = "KdBarang"
'        .TextMatrix(0, 1) = "Nama Barang"
'        .TextMatrix(0, 2) = "Asal Barang"
'        .TextMatrix(0, 3) = "Satuan"
'        .TextMatrix(0, 4) = "Stok"
'        .TextMatrix(0, 5) = "Qty"
'        .TextMatrix(0, 6) = "Harga Satuan"
'        .TextMatrix(0, 7) = "Disc"
'        .TextMatrix(0, 8) = "Total"
'        .TextMatrix(0, 9) = "KdAsal"
'        .TextMatrix(0, 10) = "TotalDiscount"
'
'        .ColWidth(0) = 1200
'        .ColWidth(1) = 3700
'        .ColWidth(2) = 1400
'        .ColWidth(3) = 800
'        .ColWidth(4) = 800
'        .ColWidth(5) = 800
'        .ColWidth(6) = 1200
'        .ColWidth(7) = 1000
'        .ColWidth(8) = 1200
'        .ColWidth(9) = 0
'        .ColWidth(10) = 0
'
'        .ColAlignment(4) = flexAlignRightCenter
'        .ColAlignment(5) = flexAlignRightCenter
'        .ColAlignment(6) = flexAlignRightCenter
'        .ColAlignment(7) = flexAlignRightCenter
'        .ColAlignment(8) = flexAlignRightCenter
'    End With
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub subLoadDcSource()
'On Error GoTo errLoad
'
'    Call msubDcSource(dcRuanganPengirim, rs, "SELECT KdRuangan, NamaRuangan FROM Ruangan ORDER BY NamaRuangan")
'    Call msubDcSource(dcAsalBarang, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where KdInstalasi = '" & mstrKdInstalasiLogin & "'")
'    If rs.EOF = False Then dcAsalBarang.BoundText = rs(0).Value
'    Call msubDcSource(dcStatusBarang, rs, "SELECT KdKelompokBarang, KelompokBarang FROM KelompokBarang ORDER BY KelompokBarang")
'    If rs.EOF = False Then dcStatusBarang.BoundText = rs(0).Value
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub subLoadText()
'Dim i As Integer
'    txtIsi.Left = fgData.Left
'    Select Case fgData.Col
'        Case 0, 1
'        Case 5
'            txtIsi.MaxLength = 4
'        Case Else
'            txtIsi.MaxLength = 0
'            Exit Sub
'    End Select
'    txtIsi.Left = fgData.Left
'
'    For i = 0 To fgData.Col - 1
'        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
'    Next i
'    txtIsi.Visible = True
'    txtIsi.Top = fgData.Top - 7
'
'    For i = 0 To fgData.Row - 1
'        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
'    Next i
'
'    If fgData.TopRow > 1 Then
'        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
'    End If
'
'    txtIsi.Width = fgData.ColWidth(fgData.Col)
''    txtIsi.Height = fgData.RowHeight(fgData.Row)
'
'    txtIsi.Visible = True
'    txtIsi.SelStart = Len(txtIsi.Text)
'    txtIsi.SetFocus
'End Sub
'
'Private Sub subLoadDataCombo(s_DcName As Object)
'Dim i As Integer
'    s_DcName.Left = fgData.Left
'    For i = 0 To fgData.Col - 1
'        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
'    Next i
'    s_DcName.Visible = True
'    s_DcName.Top = fgData.Top - 7
'
'    For i = 0 To fgData.Row - 1
'        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
'    Next i
'
'    If fgData.TopRow > 1 Then
'        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
'    End If
'
'    s_DcName.Width = fgData.ColWidth(fgData.Col)
'    s_DcName.Height = fgData.RowHeight(fgData.Row)
'
'    s_DcName.Visible = True
'    s_DcName.SetFocus
'End Sub
'
'Private Function sp_StrukKirim() As Boolean
'On Error GoTo errLoad
'    sp_StrukKirim = True
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
'        .Parameters.Append .CreateParameter("TglKirim", adDate, adParamInput, , Format(dtpTglKirim.Value, "yyyy/MM/dd HH:mm:ss"))
'        .Parameters.Append .CreateParameter("NoOrder", adChar, adParamInput, 10, IIf(substrNoOrder = "", Null, substrNoOrder))
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuanganPengirim.BoundText)
'        .Parameters.Append .CreateParameter("KdRuanganTujuan", adChar, adParamInput, 3, mstrKdRuangan)
'        .Parameters.Append .CreateParameter("IdUserPenerima", adChar, adParamInput, 10, strIDPegawaiAktif)
'        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
'        .Parameters.Append .CreateParameter("OutputNoKirim", adChar, adParamOutput, 10, Null)
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.Add_StrukKirim"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If .Parameters("return_value").Value <> 0 Then
'            MsgBox "Ada kesalahan dalam penyimpanan data struk kirim antar ruangan", vbCritical, "Validasi"
'            sp_StrukKirim = False
'        Else
'            txtNoKirim.Text = .Parameters("OutputNoKirim").Value
'        End If
'        Set dbcmd = Nothing
'        Call deleteADOCommandParameters(dbcmd)
'    End With
'Exit Function
'errLoad:
'    Call msubPesanError
'    sp_StrukKirim = False
'End Function
'
'Private Function sp_DetailBarangKeluar(f_KdBarang As String, f_KdAsal As String, f_JumlahKirim As String, _
'    f_HargaJual As Currency, f_Discount As Currency, f_PPN As Currency, f_Status As String) As Boolean
'On Error GoTo errLoad
'    sp_DetailBarangKeluar = True
'    Set dbcmd = New ADODB.Command
'    With dbcmd
'        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
'        .Parameters.Append .CreateParameter("NoKirim", adChar, adParamInput, 10, txtNoKirim.Text)
'        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, f_KdBarang)
'        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, f_KdAsal)
'        .Parameters.Append .CreateParameter("JmlKirim", adDouble, adParamInput, , CDec(f_JumlahKirim))
'        .Parameters.Append .CreateParameter("HargaJual", adCurrency, adParamInput, , f_HargaJual)
'        .Parameters.Append .CreateParameter("Discount", adDouble, adParamInput, , f_Discount)
'        .Parameters.Append .CreateParameter("Ppn", adDouble, adParamInput, , f_PPN)
'        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, dcRuanganPengirim.BoundText)
'        .Parameters.Append .CreateParameter("KdRuanganPenerima", adChar, adParamInput, 3, mstrKdRuangan)
'        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
'
'        .ActiveConnection = dbConn
'        .CommandText = "dbo.Add_DetailBarangKeluar"
'        .CommandType = adCmdStoredProc
'        .Execute
'
'        If .Parameters("return_value").Value <> 0 Then
'            MsgBox "Ada kesalahan dalam penyimpanan data detail pengiriman barang", vbCritical, "Validasi"
'            sp_DetailBarangKeluar = False
'        End If
'        Set dbcmd = Nothing
'        Call deleteADOCommandParameters(dbcmd)
'    End With
'Exit Function
'errLoad:
'    sp_DetailBarangKeluar = False
'    Call msubPesanError
'End Function
'
''Public Function subLoadDataOrder() As Boolean
''On Error GoTo errLoad
''Dim i As Integer
''
''    dgObatAlkes.Visible = False
''    Call subSetGrid
''
''    strSQL = "SELECT * FROM V_StrukOrderRuanganCetakM WHERE NoOrder = '" & txtNoOrder.Text & "' AND KdRuanganTujuan = '" & mstrKdRuangan & "'"
''    Call msubRecFO(rs, strSQL)
''
''    If rs.EOF = True Then
''        dtpTglOrder.Value = Now
''        txtRuanganPemesan.Text = ""
''        txtNamaPemesan.Text = ""
''        dcRuanganPenerima.BoundText = ""
''        txtNamaPenerima.Text = ""
''        txtKdUserPenerima.Text = ""
''        substrNoOrder = ""
''        subLoadDataOrder = False
''        Exit Function
''    End If
''
''    substrNoOrder = txtNoOrder.Text
''    subLoadDataOrder = True
''    dtpTglOrder.Value = rs("TglOrder").Value
''    txtRuanganPemesan.Text = rs("RuanganPemesan").Value
''    txtNamaPemesan.Text = rs("UserPemesan").Value
''    dcRuanganPenerima.BoundText = rs("KdRuangan").Value
''    txtNamaPenerima.Text = rs("UserPemesan").Value
''    txtKdUserPenerima.Text = rs("IdUser").Value
''    dgNamaPenerima.Visible = False
'''
''    With fgData
''        For i = 1 To rs.RecordCount
''            .TextMatrix(i, 0) = rs("KdBarang").Value
''            .TextMatrix(i, 1) = rs("Nama Barang").Value
''            .TextMatrix(i, 2) = "" 'rs("AsalBarang").Value
''            .TextMatrix(i, 3) = 0 'rs("JmlStok").Value + rs("JmlKirim").Value
''            .TextMatrix(i, 4) = IIf(IsNull(rs("JmlOrder")), 0, rs("JmlOrder"))
''            .TextMatrix(i, 5) = 0 'rs("JmlKirim").Value
''            .TextMatrix(i, 6) = 0 'rs("HargaSatuan").Value
''            'If rs("HargaSatuan") = 0 Then
''                .TextMatrix(i, 7) = 0
''            'Else
''            '    .TextMatrix(i, 7) = (rs("Discount").Value / rs("HargaSatuan").Value)
''            'End If
''            .TextMatrix(i, 8) = 0 '(rs("JmlKirim").Value * rs("HargaSatuan").Value)
''            .TextMatrix(i, 9) = 0 'rs("KdAsal").Value
''            .TextMatrix(i, 10) = 0 'rs("JmlKirim").Value * rs("Discount").Value
''            rs.MoveNext
''            .Rows = .Rows + 1
''        Next i
''        .Row = 1
''    End With
''    dgNamaPenerima.Visible = False
''Exit Function
''errLoad:
''    Call msubPesanError
''End Function
'
'Public Function subLoadDataPengiriman(f_NoTerima As String, f_KdBarang As String) As Boolean
'On Error GoTo errLoad
'Dim i As Integer
'
'    dgObatAlkes.Visible = False
'    Call subSetGrid
'    If strCetak = "ViewPengiriman" Then
'         strSQL = "SELECT * FROM V_CetakTerimaBarangLangsung WHERE NoTerima = '" & f_NoTerima & "' AND KdBarang = '" & f_KdBarang & "' AND KdRuanganTujuan= '" & mstrKdRuangan & "'"
'         strCetak = ""
'         'cmdSimpan.Enabled = False
'         'cmdBatal.Enabled = False
'    Else
'        strSQL = "SELECT * FROM V_CetakTerimaBarangLangsung WHERE NoTerima = '" & f_NoTerima & "' AND KdBarang = '" & f_KdBarang & "' AND KdRuangan= '" & mstrKdRuangan & "'"
'    End If
'
'        Call msubRecFO(rs, strSQL)
'
'    If rs.EOF = True Then
'        'dtpTglOrder.Value = Now
'        'txtNamaPemesan.Text = ""
'        'dcRuanganPenerima.BoundText = ""
'        substrNoOrder = ""
'        subLoadDataPengiriman = False
'        Exit Function
'    End If
'
''    txtNoOrder.Text = IIf(IsNull(rs("NoOrder")), "", rs("NoOrder"))
''    substrNoOrder = txtNoOrder.Text
'    subLoadDataPengiriman = True
''    txtNamaPemesan.Text = IIf(IsNull(rs("UserPemesan")), "", rs("UserPemesan"))
'    txtNoKirim.Text = f_NoTerima
'    dtpTglKirim.Value = rs("TglTerima")
'    'dcRuanganPenerima.BoundText = rs("KdRuanganTujuan").Value
'    'txtNamaPenerima.Text = rs("UserPenerima").Value
'    'txtKdUserPenerima.Text = rs("IdUserPenerima").Value
'    dgNamaPenerima.Visible = False
'
'    With fgData
'        For i = 1 To rs.RecordCount
'            .TextMatrix(i, 0) = rs("KdBarang").Value
'            .TextMatrix(i, 1) = rs("Nama Barang").Value
'            .TextMatrix(i, 2) = rs("AsalBarang").Value
'            .TextMatrix(i, 3) = rs("Satuan").Value
'            .TextMatrix(i, 4) = IIf(IsNull(rs("JmlStok") + rs("Jumlah")), 0, rs("JmlStok") + rs("Jumlah"))
'            .TextMatrix(i, 5) = IIf(IsNull(rs("Jumlah")), 0, rs("Jumlah"))
'            .TextMatrix(i, 6) = rs("HargaSatuan").Value
'            If rs("HargaSatuan") = 0 Then
'                .TextMatrix(i, 7) = 0
'            Else
'                .TextMatrix(i, 7) = (rs("Discount").Value / rs("HargaSatuan").Value)
'            End If
'            .TextMatrix(i, 8) = (rs("Jumlah").Value * rs("HargaSatuan").Value)
'            .TextMatrix(i, 9) = rs("KdAsal").Value
'            .TextMatrix(i, 10) = rs("Jumlah").Value * rs("Discount").Value
'            rs.MoveNext
'            .Rows = .Rows + 1
'        Next i
'        .Row = 1
'    End With
'    Call subHitungTotal
'
'    Call msubRecFO(dbRst, "SELECT TglTerima, Ruangan, Operator FROM  V_CetakTerimaBarangLangsung WHERE NoTerima = '" & txtNoKirim.Text & "'")
'    If dbRst.EOF = True Then
'        dtpTglKirim.Value = Now
'       ' txtRuanganPemesan.Text = ""
'        'txtNamaPemesan.Text = ""
'    Else
'        dtpTglKirim.Value = dbRst("TglTerima")
'        dcRuanganPengirim.Text = dbRst("Ruangan")
'        'txtNamaPemesan.Text = dbRst("Operator")
'    End If
'
'Exit Function
'errLoad:
'    Call msubPesanError
'End Function
'
'Private Sub txtNamaPenerima_Change()
'On Error GoTo errLoad
'Dim i As Integer
'
'    strSQL = " SELECT [Nama Pemeriksa], JK, [Jenis Pemeriksa], IdPegawai " & _
'        " From V_DaftarPemeriksaPasien" & _
'        " where [Nama Pemeriksa] like '" & txtNamaPenerima.Text & "%' " & _
'        " ORDER BY [Nama Pemeriksa], [Jenis Pemeriksa]"
'    Call msubRecFO(dbRst, strSQL)
'
'    Set dgNamaPenerima.DataSource = dbRst
'    With dgNamaPenerima
'        .Columns("Nama Pemeriksa").Width = 2000
'        .Columns("JK").Width = 360
'        .Columns("Jenis Pemeriksa").Width = 1500
'        .Columns("IdPegawai").Width = 0
'
'        .Columns("JK").Alignment = dbgCenter
'    End With
'    dgNamaPenerima.Visible = True
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'
'Private Sub txtNamaPenerima_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case 13
'            If dgNamaPenerima.Visible = True Then
'                dgNamaPenerima.SetFocus
'            Else
'                fgData.Col = 1
'                fgData.SetFocus
'            End If
'        Case 27
'            dgNamaPenerima.Visible = False
'    End Select
'End Sub
'
''Private Sub txtNoOrder_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        If subLoadDataOrder = True Then dtpTglKirim.SetFocus
''        Call subHitungTotal
''    End If
''End Sub
'
'Private Sub subHitungTotal()
'On Error GoTo errLoad
'Dim i As Integer
'
'    txtTotalBiaya.Text = 0
'    txtTotalDiscount.Text = 0
'
'    With fgData
'        For i = 1 To fgData.Rows - 1
'            If .TextMatrix(i, 8) = "" Then .TextMatrix(i, 8) = 0
'            If .TextMatrix(i, 10) = "" Then .TextMatrix(i, 10) = 0
'            txtTotalBiaya.Text = txtTotalBiaya.Text + Val(.TextMatrix(i, 8))
'            txtTotalDiscount.Text = txtTotalDiscount.Text + Val(.TextMatrix(i, 10))
'        Next i
'    End With
'
'    txtTotalBiaya.Text = IIf(Val(txtTotalBiaya.Text) = 0, 0, Format(txtTotalBiaya.Text, "#,###"))
'    txtTotalDiscount.Text = IIf(Val(txtTotalDiscount.Text) = 0, 0, Format(txtTotalDiscount.Text, "#,###"))
'
'Exit Sub
'errLoad:
'    Call msubPesanError
'End Sub
'