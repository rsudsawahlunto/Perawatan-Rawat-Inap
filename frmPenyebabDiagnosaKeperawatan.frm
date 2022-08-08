VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPenyebabDiagnosaKeperawatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Penyebab Diagnosa Keperawatan"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPenyebabDiagnosaKeperawatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7485
   Begin VB.TextBox txtOuputKode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   7455
      Begin VB.CheckBox CheckStatusEnbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   5880
         TabIndex        =   16
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaExternal1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   5535
      End
      Begin VB.TextBox txtKodeExternal1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtPenyebab 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox txtKode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "123"
         Top             =   600
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   3615
         Left            =   240
         TabIndex        =   1
         Top             =   2520
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   6376
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
      Begin VB.Label Label8 
         Caption         =   "Nama External"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Kode External"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Penyebab Diagnosa Keperawatan"
         Height          =   210
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   2730
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   7200
      Width           =   7455
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   6045
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   3615
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4830
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F1 - Cetak"
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   930
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   13
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
      Picture         =   "frmPenyebabDiagnosaKeperawatan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5640
      Picture         =   "frmPenyebabDiagnosaKeperawatan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPenyebabDiagnosaKeperawatan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmPenyebabDiagnosaKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilterDiagnosa As String
Dim intJmlDetail As Integer
Dim mstrKdDetail As String

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call clear
    Call subLoadGridSource
    txtPenyebab.SetFocus
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errLoad
    If txtKode.Text = "" Then
        MsgBox "Pilih dulu Detail Askep yang akan dihapus", vbOKOnly, "Validasi"
        Exit Sub
    End If
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    If sp_PenyebabAskep("D") = False Then Exit Sub

    Call cmdBatal_Click
errLoad:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Periksa("text", txtPenyebab, "Penyebab Askep kosong") = False Then Exit Sub
    If sp_PenyebabAskep("A") = False Then Exit Sub
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Call cmdBatal_Click
errLoad:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgData_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgData
    WheelHook.WheelHook dgData
End Sub

Private Sub dgData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPenyebab.SetFocus
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad

    If dgData.ApproxCount = 0 Then Exit Sub
    txtKode.Text = dgData.Columns(0).Value
    txtPenyebab.Text = dgData.Columns(1).Value
    txtKodeExternal1.Text = dgData.Columns(2).Value
    txtNamaExternal1.Text = dgData.Columns(3).Value
    If dgData.Columns(4) = "" Then
        CheckStatusEnbl1.Value = 0
    ElseIf dgData.Columns(4) = 0 Then
        CheckStatusEnbl1.Value = 0
    ElseIf dgData.Columns(4) = 1 Then
        CheckStatusEnbl1.Value = 1
    End If

    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If dgData.ApproxCount = 0 Then Exit Sub
            frmCetakPenyebabDiagnosaKeperawatan.Show
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call openConnection
    Call clear
    Call subLoadGridSource

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
    Set rs = Nothing
    strSQL = "select * from PenyebabDiagnosaKeperawatan"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgData.DataSource = rs
    With dgData
        .Columns(0).Caption = "Kode"
        .Columns(0).Width = 1500
        .Columns(1).Width = 4800
    End With
    Set rs = Nothing
End Sub

Private Sub clear()
    txtKode.Text = ""
    txtPenyebab.Text = ""
    txtKodeExternal1.Text = ""
    txtNamaExternal1.Text = ""
    CheckStatusEnbl1.Value = 1
End Sub

Private Sub txtPenyebab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKodeExternal1.SetFocus
End Sub

Private Sub txtPenyebab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Function sp_PenyebabAskep(f_Status As String) As Boolean
    sp_PenyebabAskep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPenyebab", adVarChar, adParamInput, 3, txtKode.Text)
        .Parameters.Append .CreateParameter("PenyebabAskep", adVarChar, adParamInput, 50, Trim(txtPenyebab.Text))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_PenyebabDiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_PenyebabAskep = False
        Else
            Call Add_HistoryLoginActivity("AUD_PenyebabDiagnosaKeperawatan")
        End If

        If f_Status = "A" Then
            txtOuputKode.Text = .Parameters("OutputKode").Value
        Else
            MsgBox "Berhasil menghapus data", vbInformation, "Informasi"
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtKodeExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal1.SetFocus
End Sub

Private Sub txtNamaExternal1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CheckStatusEnbl1.SetFocus
End Sub

Private Sub CheckStatusEnbl1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
