VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataDiagnosa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Diagnosa Ruangan"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8535
   Begin VB.CheckBox chkAll 
      Caption         =   "Pilih Semua"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame framDiagnosa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   8535
      Begin VB.TextBox txtDiagnosa 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3360
         TabIndex        =   1
         Top             =   480
         Width           =   4935
      End
      Begin MSDataListLib.DataCombo dcSubInstalasi 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SMF (Kasus Penyakit)"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diagnosa"
         Height          =   210
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1230
      End
   End
   Begin MSComctlLib.ListView lvwDiagnosa 
      Height          =   4695
      Left            =   5
      TabIndex        =   2
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Diagnosa"
         Object.Width           =   13229
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   6720
      Picture         =   "frmDataDiagnosa.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataDiagnosa.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label lblBanyakData 
      AutoSize        =   -1  'True
      Caption         =   "10 / 1000 diagnosa"
      Height          =   210
      Left            =   0
      TabIndex        =   9
      Top             =   6840
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataDiagnosa.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDataDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim i As Integer

Private Sub chkAll_Click()
    If chkAll.Value = vbChecked Then
        For i = 1 To lvwDiagnosa.ListItems.Count
            lvwDiagnosa.ListItems(i).Checked = True
        Next i
    Else
        For i = 1 To lvwDiagnosa.ListItems.Count
            lvwDiagnosa.ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan

    If dcSubInstalasi.Text = "" Then
        MsgBox "Pilihan SubInstalasi harus diisi", vbCritical, "Validasi"
        Exit Sub
    End If
    cmdSimpan.Enabled = False
    For i = 1 To lvwDiagnosa.ListItems.Count

        If lvwDiagnosa.ListItems(i).Checked = True Then
            If sp_DiagnosaRuangan(lvwDiagnosa.ListItems(i).Key, "A") = False Then Exit Sub
        Else
            If sp_DiagnosaRuangan(lvwDiagnosa.ListItems(i).Key, "D") = False Then Exit Sub
        End If
    Next i
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call Add_HistoryLoginActivity("AUD_DiagnosaRuangan")
    cmdSimpan.Enabled = True
    Exit Sub
errSimpan:
    MsgBox "Ada kesalahan, hubungi administrator & laporkan pesan error berikut" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcSubInstalasi_Change()
    Call txtDiagnosa_Change
End Sub

Private Sub dcSubInstalasi_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcSubInstalasi.Text)) = 0 Then txtDiagnosa.SetFocus: Exit Sub
        If dcSubInstalasi.MatchedWithList = True Then txtDiagnosa.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT SubInstalasiRuangan.KdSubInstalasi,SubInstalasi.NamaSubInstalasi FROM SubInstalasiRuangan INNER JOIN SubInstalasi ON SubInstalasiRuangan.KdSubInstalasi=SubInstalasi.KdSubInstalasi WHERE NamaSubInstalasi LIKE '%" & dcSubInstalasi.Text & "%' ")
        If dbRst.EOF = True Then Exit Sub
        dcSubInstalasi.BoundText = dbRst(0).Value
        dcSubInstalasi.Text = dbRst(1).Value
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    strSQL = "SELECT SubInstalasiRuangan.KdSubInstalasi," _
    & "SubInstalasi.NamaSubInstalasi FROM SubInstalasiRuangan INNER JOIN " _
    & "SubInstalasi ON SubInstalasiRuangan.KdSubInstalasi=" _
    & "SubInstalasi.KdSubInstalasi WHERE SubInstalasiRuangan.KdRuangan='" _
    & mstrKdRuangan & "' and SubInstalasi.StatusEnabled='1' order by SubInstalasi.NamaSubInstalasi"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data Sub Instalasi", vbInformation, ""
    Else
        Set dcSubInstalasi.RowSource = rs
        dcSubInstalasi.BoundColumn = rs(0).Name
        dcSubInstalasi.ListField = rs(1).Name
        dcSubInstalasi.BoundText = rs(0).Value
        Set rs = Nothing
    End If
    Call txtDiagnosa_Change
    Call PlayFlashMovie(Me)

    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub txtDiagnosa_Change()
    On Error GoTo errLoad
    Dim intJumlahDiagnosa As Integer
    Dim j As Integer

    strSQL = "SELECT top 200 KdDiagnosa,NamaDiagnosa FROM Diagnosa WHERE NamaDiagnosa LIKE '%" & txtDiagnosa.Text & "%' ORDER BY NamaDiagnosa ASC"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

    lvwDiagnosa.ListItems.clear
    intJumlahDiagnosa = 0

    Do While rs.EOF = False
        Set itemAll = lvwDiagnosa.ListItems.Add(, rs(0).Value, rs(1).Value)
        rs.MoveNext
    Loop
    strSQL = "SELECT KdDiagnosa FROM DiagnosaRuangan WHERE KdSubInstalasi='" & dcSubInstalasi.BoundText & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly
    If rs.EOF = True Then Exit Sub
    For i = 1 To lvwDiagnosa.ListItems.Count
        rs.MoveFirst
        While Not rs.EOF
            If lvwDiagnosa.ListItems(i).Key = rs(0).Value Then
                lvwDiagnosa.ListItems(i).Checked = True
                lvwDiagnosa.ListItems(i).ForeColor = vbBlue
                intJumlahDiagnosa = intJumlahDiagnosa + 1
            End If
            rs.MoveNext

        Wend

    Next i
    lblBanyakData.Caption = intJumlahDiagnosa & " / " & lvwDiagnosa.ListItems.Count & " diagnosa"
    Exit Sub
errLoad:
    MsgBox "Ada kesalahan, hubungi administrator & laporkan pesan error berikut" _
    & vbNewLine & Err.Number & " - " & Err.Description, vbCritical, "Validasi"
End Sub

Private Function sp_DiagnosaRuangan(f_KdDiagnosa As String, f_status As String) As Boolean
    On Error GoTo hell
    sp_DiagnosaRuangan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDiagnosa", adVarChar, adParamInput, 7, f_KdDiagnosa)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adVarChar, adParamInput, 3, dcSubInstalasi.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        Set .ActiveConnection = dbConn
        .CommandText = "AUD_DiagnosaRuangan"
        .CommandType = adCmdStoredProc
        .Execute

        If (.Parameters("return_Value").Value) <> 0 Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            sp_DiagnosaRuangan = False
            Set dbcmd = Nothing
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
hell:
    sp_DiagnosaRuangan = False
    Call msubPesanError
End Function

Private Sub txtDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
