VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form frmStokBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Stock Barang"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStokBarang.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9975
   Begin MSDataGridLib.DataGrid dgCariBarang 
      Height          =   2535
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
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
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   9735
      Begin VB.TextBox txtCariBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4560
         Width           =   3240
      End
      Begin MSDataGridLib.DataGrid dgStockBarang 
         Height          =   4095
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   15
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
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari Barang"
         Height          =   210
         Index           =   6
         Left            =   255
         TabIndex        =   20
         Top             =   4605
         Width           =   900
      End
      Begin VB.Label lblJmlData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Barang"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8160
         TabIndex        =   19
         Top             =   4620
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   6510
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4935
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8085
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame fraBarang 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   9735
      Begin VB.TextBox txtLokasi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8160
         MaxLength       =   25
         TabIndex        =   4
         Top             =   600
         Width           =   1305
      End
      Begin VB.TextBox txtKdBarang 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "txtkdbarang"
         Top             =   240
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.TextBox txtNamaBarang 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   3120
      End
      Begin VB.TextBox txtJmlMinimum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5520
         MaxLength       =   25
         TabIndex        =   2
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtJmlStock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6840
         MaxLength       =   25
         TabIndex        =   3
         Top             =   600
         Width           =   1200
      End
      Begin MSDataListLib.DataCombo dcAsalBarang 
         Height          =   330
         Left            =   3480
         TabIndex        =   1
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Lokasi"
         Height          =   210
         Index           =   4
         Left            =   8160
         TabIndex        =   21
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Nama Barang"
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Asal Barang"
         Height          =   210
         Index           =   1
         Left            =   3480
         TabIndex        =   16
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jml Minimum"
         Height          =   210
         Index           =   2
         Left            =   5520
         TabIndex        =   15
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jml Stock"
         Height          =   210
         Index           =   3
         Left            =   6840
         TabIndex        =   14
         Top             =   360
         Width           =   780
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
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
      Left            =   8160
      Picture         =   "frmStokBarang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStokBarang.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStokBarang.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmStokBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
    On Error GoTo errLoad

    Call subKosong
    Call subLoadDcSource
    Call subLoadGridSource
    txtNamaBarang.SetFocus

    Exit Sub
errLoad:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad

    If txtKdBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If
    If Periksa("datacombo", dcAsalBarang, "Asal barang kosong") = False Then Exit Sub

    If MsgBox("Anda yakin akan menghapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    dbConn.Execute "DELETE StokRuangan WHERE KdBarang ='" & txtKdBarang.Text & "' AND KdAsal='" & dcAsalBarang.BoundText & "' AND KdRuangan='" & mstrKdRuangan & "'"
    Call cmdBatal_Click
    MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    If txtKdBarang.Text = "" Then
        MsgBox "Nama barang kosong", vbExclamation, "Validasi": txtNamaBarang.SetFocus: Exit Sub
    End If
    If Periksa("datacombo", dcAsalBarang, "Asal barang kosong") = False Then Exit Sub

    If sp_StockBarang() = False Then Exit Sub
    Call cmdBatal_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcAsalBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlMinimum.SetFocus
End Sub

Private Sub dgCariBarang_DblClick()
    On Error GoTo errLoad

    With dgCariBarang
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("Nama Barang")
        .Visible = False
    End With
    dcAsalBarang.SetFocus

    Exit Sub
errLoad:
End Sub

Private Sub dgCariBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgCariBarang_DblClick
End Sub

Private Sub dgStockBarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaBarang.SetFocus
End Sub

Private Sub dgStockBarang_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    With dgStockBarang
        If .ApproxCount = 0 Then Exit Sub
        txtKdBarang.Text = .Columns("KdBarang")
        txtNamaBarang.Text = .Columns("Nama Barang")
        dcAsalBarang.BoundText = .Columns("KdAsal")
        txtJmlMinimum.Text = .Columns("Jml. Min")
        txtJmlStock.Text = .Columns("Jml. Stock")
        txtLokasi.Text = .Columns("Lokasi")
    End With
    dgCariBarang.Visible = False
    lblJmlData.Caption = dgStockBarang.Bookmark & " / " & dgStockBarang.ApproxCount & " Data"

    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call cmdBatal_Click
End Sub

Private Sub txtCariBarang_Change()
    On Error GoTo errLoad

    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtJmlMinimum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJmlStock.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtJmlMinimum_LostFocus()
    txtJmlMinimum.Text = IIf(val(txtJmlMinimum) = 0, 0, FormatPembulatan(CDbl(txtJmlMinimum), mstrKdInstalasiLogin))
End Sub

Private Sub txtJmlStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLokasi.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = 44) Then KeyAscii = 0
End Sub

Private Sub txtJmlStock_LostFocus()
    txtJmlStock.Text = IIf(val(txtJmlStock) = 0, 0, Format(txtJmlStock, "#,##0.##"))
End Sub

Private Sub txtLokasi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then dgStockBarang.SetFocus
End Sub

Private Sub txtLokasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaBarang_Change()
    On Error GoTo errLoad

    Call subCariBarang

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaBarang_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDown Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus
    If KeyCode = vbKeyEscape Then dgCariBarang.Visible = False
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then If dgCariBarang.Visible = True Then dgCariBarang.SetFocus Else dcAsalBarang.SetFocus
End Sub

Private Sub subKosong()
    txtKdBarang.Text = ""
    txtNamaBarang.Text = ""
    dcAsalBarang.BoundText = ""
    txtJmlMinimum.Text = 0
    txtJmlStock.Text = 0
    txtLokasi.Text = ""
    dgCariBarang.Visible = False
End Sub

Private Sub subLoadDcSource()
    On Error GoTo errLoad

    Call msubDcSource(dcAsalBarang, rs, "SELECT KdAsal, NamaAsal FROM AsalBarang where KdInstalasi = '" & mstrKdInstalasiLogin & "' ORDER BY NamaAsal")

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subCariBarang()
    On Error GoTo errLoad

    strSQL = "SELECT  [Nama Barang], Kekuatan, Satuan, DetailJenisBrg AS [Jenis Barang], KdBarang FROM V_CariBarang " & _
    " WHERE [Nama Barang] LIKE '%" & txtNamaBarang.Text & "%'" & _
    " ORDER BY [Nama Barang]"
    Call msubRecFO(rs, strSQL)
    Set dgCariBarang.DataSource = rs
    With dgCariBarang
        .Columns("Nama Barang").Width = 2900
        .Columns("Kekuatan").Width = 1000
        .Columns("Satuan").Width = 1000
        .Columns("Jenis Barang").Width = 1440
        .Columns("KdBarang").Width = 0
        .Height = 2390
        .Visible = True
    End With

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    On Error GoTo errLoad
    Dim i As Integer

    strSQL = "SELECT TOP 100 NamaBarang AS [Nama Barang], [Asal Barang] AS NamaAsal, JmlMinimum AS [Jml. Min], JmlStok AS [Jml. Stock], [Jenis Barang], KdBarang, KdAsal, KdDetailJenisBarang, KdRuangan, Lokasi " & _
    " FROM V_AmbilStockBarang " & _
    " WHERE [Nama Barang] LIKE '%" & txtCariBarang & "%' AND KdRuangan = '" & mstrKdRuangan & "'"
    Call msubRecFO(rs, strSQL)
    Set dgStockBarang.DataSource = rs
    With dgStockBarang
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next i
        .Columns("Nama Barang").Width = 3300
        .Columns("NamaAsal").Width = 1000
        .Columns("Jml. Min").Width = 1000
        .Columns("Jml. Stock").Width = 1000
        .Columns("Jml. Stock").NumberFormat = "#,##0.00"
        .Columns("Jenis Barang").Width = 1200
        .Columns("Lokasi").Width = 1000
        .Columns("Jml. Stock").Alignment = dbgRight
        .Columns("Lokasi").Alignment = dbgRight
    End With
    lblJmlData.Caption = 0 & " / " & dgStockBarang.ApproxCount & " Data"

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Function sp_StockBarang() As Boolean
    On Error GoTo errLoad

    sp_StockBarang = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, txtKdBarang.Text)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, dcAsalBarang.BoundText)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("JmlMin", adDouble, adParamInput, , CDbl(txtJmlMinimum.Text))
        .Parameters.Append .CreateParameter("JmlStok", adDouble, adParamInput, , CDec(txtJmlStock.Text))
        .Parameters.Append .CreateParameter("Lokasi", adVarChar, adParamInput, 12, IIf(txtLokasi.Text = "", Null, txtLokasi.Text))

        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_StokBarangRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_StockBarang = False
        Else
            Call Add_HistoryLoginActivity("AU_StokBarangRuangan")
        End If
    End With
    Set dbcmd = Nothing
    Call deleteADOCommandParameters(dbcmd)
    Exit Function
errLoad:
    Call msubPesanError
End Function
