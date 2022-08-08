VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMasterVKbersalin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Jenis Persalinan & Event Bayi "
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmMasterVKbersalin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   7365
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   7440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Persalinan"
      TabPicture(0)   =   "frmMasterVKbersalin.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "dgJenisPersalinan"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Event Bayi"
      TabPicture(1)   =   "frmMasterVKbersalin.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dgEventBayi"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkSts 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5520
            TabIndex        =   3
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1320
            Width           =   4695
         End
         Begin VB.TextBox txtKdExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   2
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtJenisPersalinan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   1
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txtKdJenisPersalinan 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Persalinan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis Persalinan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1725
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkSts1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5520
            TabIndex        =   13
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExt1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   14
            Top             =   1320
            Width           =   5175
         End
         Begin VB.TextBox txtKdExt1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   12
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtKodeEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtNamaEvent 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   11
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Kode Event"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nama Event"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   990
         End
      End
      Begin MSDataGridLib.DataGrid dgJenisPersalinan 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   5
         Top             =   2400
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
               LCID            =   1033
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
               LCID            =   1033
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
      Begin MSDataGridLib.DataGrid dgEventBayi 
         Height          =   3735
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
               LCID            =   1033
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
               LCID            =   1033
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
      Left            =   5520
      Picture         =   "frmMasterVKbersalin.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterVKbersalin.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterVKbersalin.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterVKbersalin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBatal_Click()
On Error Resume Next
    Call clearData
End Sub

Public Sub clearData()
On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            txtKdJenisPersalinan.Text = ""
            txtJenisPersalinan.Text = ""
            txtJenisPersalinan.SetFocus
            txtKdExt.Text = ""
            txtNmExt.Text = ""
            chkSts.Value = 1
        Case 1
            txtKodeEvent.Text = ""
            txtNamaEvent.Text = ""
            txtNamaEvent.SetFocus
            txtKdExt1.Text = ""
            txtNmExt1.Text = ""
            chkSts1.Value = 1
    End Select
    Call loadGridSource
End Sub

Public Sub loadGridSource()
On Error GoTo errLoad
    Set rs = Nothing
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT * FROM JenisPersalinan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With dgJenisPersalinan
                Set .DataSource = rs
                .Columns(0).Width = 1000
                .Columns(1).Width = 3500
            End With
        Case 1
            strSQL = "SELECT * FROM EventBayi"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            With dgEventBayi
                Set .DataSource = rs
                .Columns(0).Width = 1000
                .Columns(1).Width = 3500
            End With
    End Select
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub cmdHapus_Click()
On Error GoTo hell
    Select Case SSTab1.Tab
        Case 0
            If txtKdJenisPersalinan.Text = "" Then Exit Sub
            If MsgBox("Yakin akan menghapus data Jenis Persalinan", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If AUD_JenisPersalinan("D") = False Then Exit Sub
            
        Case 1
            If txtKodeEvent.Text = "" Then Exit Sub
            If MsgBox("Yakin akan menghapus data Nama Event Bayi", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            If AUD_EventBayi("D") = False Then Exit Sub
            
    End Select
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdBatal_Click
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo hell
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisPersalinan, "Jenis Persalinan kosong") = False Then Exit Sub
            If AUD_JenisPersalinan("A") = False Then Exit Sub
        Case 1
            If Periksa("text", txtNamaEvent, "Nama event Kosong") = False Then Exit Sub
            If AUD_EventBayi("A") = False Then Exit Sub
            
    End Select
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdBatal_Click
    cmdTutup.SetFocus
Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
Private Function AUD_JenisPersalinan(f_status As String) As Boolean
On Error GoTo hell_
Set dbcmd = New ADODB.Command
    AUD_JenisPersalinan = True
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPersalinan", adChar, adParamInput, 2, txtKdJenisPersalinan.Text)
        .Parameters.Append .CreateParameter("JenisPersalinan", adVarChar, adParamInput, 75, txtJenisPersalinan.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(Trim(txtKdExt.Text = ""), Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 75, IIf(Trim(txtNmExt.Text = ""), Null, Trim(txtNmExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
                
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_JenisPersalinan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            AUD_JenisPersalinan = False
        Else
            Call Add_HistoryLoginActivity("AUD_JenisPersalinan")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell_:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    msubPesanError
    AUD_JenisPersalinan = False
End Function
Private Function AUD_EventBayi(f_status As String) As Boolean
On Error GoTo hell_
Set dbcmd = New ADODB.Command
    AUD_EventBayi = True
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdEvent", adChar, adParamInput, 3, txtKodeEvent.Text)
        .Parameters.Append .CreateParameter("NamaEvent", adVarChar, adParamInput, 75, txtNamaEvent.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(Trim(txtKdExt1.Text = ""), Null, Trim(txtKdExt1.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 75, IIf(Trim(txtNmExt1.Text = ""), Null, Trim(txtNmExt1.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
                
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_EventBayiLahir"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan data", vbCritical, "Validasi"
            AUD_EventBayi = False
        Else
            Call Add_HistoryLoginActivity("AUD_EventBayiLahir")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
hell_:
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    msubPesanError
    AUD_EventBayi = False
End Function

Private Sub dgEventBayi_Click()
        WheelHook.WheelUnHook
        Set MyProperty = dgEventBayi
        WheelHook.WheelHook dgEventBayi
End Sub

Private Sub dgEventBayi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    txtKodeEvent.Text = dgEventBayi.Columns("KdEvent")
    txtNamaEvent.Text = dgEventBayi.Columns("NamaEvent")
'    txtKdExt1.Text = dgEventBayi.Columns("KodeExternal")
'    txtNmExt1.Text = dgEventBayi.Columns("NamaExternal")
    If dgEventBayi.Columns("KodeExternal").Value = "" Then txtKdExt1.Text = "" Else txtKdExt1.Text = dgEventBayi.Columns("KodeExternal").Value
    If dgEventBayi.Columns("NamaExternal").Value = "" Then txtNmExt1.Text = "" Else txtNmExt1.Text = dgEventBayi.Columns("NamaExternal").Value
    chkSts1.Value = dgEventBayi.Columns("StatusEnabled").Value
End Sub

Private Sub dgJenisPersalinan_Click()
        WheelHook.WheelUnHook
        Set MyProperty = dgJenisPersalinan
        WheelHook.WheelHook dgJenisPersalinan
End Sub

Private Sub dgJenisPersalinan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    txtKdJenisPersalinan.Text = dgJenisPersalinan.Columns("KdJenisPersalinan")
    txtJenisPersalinan.Text = dgJenisPersalinan.Columns("JenisPersalinan")
'    txtKdExt.Text = dgJenisPersalinan.Columns("KodeExternal").Value
'    txtNmExt.Text = dgJenisPersalinan.Columns("NamaExternal").Value
    If dgJenisPersalinan.Columns("KodeExternal").Value = "" Then txtKdExt.Text = "" Else txtKdExt.Text = dgJenisPersalinan.Columns("KodeExternal").Value
    If dgJenisPersalinan.Columns("NamaExternal").Value = "" Then txtNmExt.Text = "" Else txtNmExt.Text = dgJenisPersalinan.Columns("NamaExternal").Value
    chkSts.Value = dgJenisPersalinan.Columns("StatusEnabled").Value
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    SSTab1.Tab = 0
    Call loadGridSource
    Call cmdBatal_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo errLoad
    Call loadGridSource
Exit Sub
errLoad:
End Sub

Private Sub txtJenisPersalinan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtNamaEvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt1.SetFocus
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkSts.SetFocus
End Sub

Private Sub chkSts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt.SetFocus
End Sub

Private Sub txtNmExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtKdExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkSts1.SetFocus
End Sub

Private Sub chkSts1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt1.SetFocus
End Sub

Private Sub txtNmExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

