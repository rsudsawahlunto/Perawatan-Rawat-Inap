VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterDiagnosaAskep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Diagnosa Asuhan Keperawatan"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterDiagnosaAskep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10485
   Begin VB.TextBox txtOutputKode2 
      Height          =   375
      Left            =   0
      TabIndex        =   34
      Text            =   "12345"
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOuputKode 
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Text            =   "123"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
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
      TabIndex        =   2
      Top             =   8640
      Width           =   10455
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4847
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   3615
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   6045
         TabIndex        =   3
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
         TabIndex        =   7
         Top             =   360
         Width           =   930
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Diagnosa Keperawatan"
      TabPicture(0)   =   "frmMasterDiagnosaAskep.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detail Diagnosa Keperawatan"
      TabPicture(1)   =   "frmMasterDiagnosaAskep.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtKdDetailAskep 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   3
            TabIndex        =   26
            Text            =   "123"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtDetailAskep 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   25
            Top             =   600
            Width           =   6135
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   7920
            TabIndex        =   24
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2280
            TabIndex        =   23
            Top             =   1200
            Width           =   5535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid dgDetailAskep 
            Height          =   3615
            Left            =   240
            TabIndex        =   27
            Top             =   1680
            Width           =   9480
            _ExtentX        =   16722
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
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Detail Asuhan Keperawatan"
            Height          =   210
            Left            =   1080
            TabIndex        =   30
            Top             =   360
            Width           =   2250
         End
         Begin VB.Label Label7 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   960
            Width           =   1335
         End
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
         Height          =   6855
         Left            =   -74760
         TabIndex        =   8
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtDiagnosa 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1155
            Left            =   240
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   480
            Width           =   9615
         End
         Begin VB.TextBox txtKdDiagnosaKeperawatan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   240
            MaxLength       =   5
            TabIndex        =   33
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame fraDiagnosa 
            Caption         =   "Data Diagnosa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   240
            TabIndex        =   17
            Top             =   2880
            Visible         =   0   'False
            Width           =   9615
            Begin MSDataGridLib.DataGrid dgDiagnosa 
               Height          =   2775
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   9375
               _ExtentX        =   16536
               _ExtentY        =   4895
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
         End
         Begin VB.TextBox txtKodeExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   11
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox txtNamaExternal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Top             =   2280
            Width           =   6015
         End
         Begin VB.CheckBox CheckStatusEnbl 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   8280
            TabIndex        =   9
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dgDiagnosaKeperawatan 
            Height          =   3855
            Left            =   240
            TabIndex        =   12
            Top             =   2760
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   6800
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
         Begin MSDataListLib.DataCombo dcAskep 
            Height          =   330
            Left            =   2160
            TabIndex        =   20
            Top             =   1680
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label4 
            Caption         =   "Asuhan Keperawatan"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1755
            Width           =   2775
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Diagnosa Keperawatan"
            Height          =   210
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1860
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kode Diagnosa"
            Height          =   210
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label2 
            Caption         =   "Kode External"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Nama External"
            Height          =   255
            Left            =   2160
            TabIndex        =   13
            Top             =   2040
            Width           =   1335
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8640
      Picture         =   "frmMasterDiagnosaAskep.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDiagnosaAskep.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterDiagnosaAskep.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterDiagnosaAskep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilterDiagnosa As String
Dim intJmlDiagnosa As Integer
Dim mstrKdDiagnosa As String

Private Sub cmdBatal_Click()
    Call clear
    Call subLoadGridSource
   ' txtKdDiagnosa.SetFocus
    txtDiagnosa.SetFocus
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errLoad
 If SSTab1.Tab = 0 Then
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    If sp_DiagnosaKeperawatan("D") = False Then Exit Sub

    Call cmdBatal_Click
 ElseIf SSTab1.Tab = 1 Then
    
    On Error GoTo errLoad
    If txtKdDetailAskep.Text = "" Then
        MsgBox "Pilih dulu Detail Askep yang akan dihapus", vbOKOnly, "Validasi"
        Exit Sub
    End If
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    If sp_DetailAskep("D") = False Then Exit Sub

    Call cmdBatal_Click
 
 End If
 
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
  If SSTab1.Tab = 0 Then
    'If Periksa("text", txtKdDiagnosa, "Kode Diagnosa kosong") = False Then Exit Sub
    If Periksa("text", txtDiagnosa, "Diagnosa Keperawatan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcAskep, "Asuhan Keperawatan kosong") = False Then Exit Sub
    If sp_DiagnosaKeperawatan("A") = False Then Exit Sub
    
    Call cmdBatal_Click
  ElseIf SSTab1.Tab = 1 Then
    On Error GoTo errLoad
    If Periksa("text", txtDetailAskep, "Detail Askep kosong") = False Then Exit Sub
    If sp_DetailAskep("A") = False Then Exit Sub
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Call cmdBatal_Click
  End If
  
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdCetak_Click()
    frmCetakMastDiagnosaKeperawatan.Show
End Sub

Private Sub dgDiagnosa_Click()
'    Call dgDiagnosa_KeyPress(13)
WheelHook.WheelUnHook
        Set MyProperty = dgDiagnosa
        WheelHook.WheelHook dgDiagnosa
End Sub

Private Sub dgDiagnosa_DblClick()
    Call dgDiagnosa_KeyPress(13)
End Sub

Private Sub dgDiagnosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlDiagnosa = 0 Then Exit Sub
        mstrKdDiagnosa = dgDiagnosa.Columns(0).Value
        If mstrKdDiagnosa = "" Then
            MsgBox "Pilih dulu Diagnosa-nya", vbCritical, "Validasi"
            dgDiagnosa.SetFocus
            Exit Sub
        End If
        fraDiagnosa.Visible = False
        txtDiagnosa.SetFocus
    End If
    If KeyAscii = 27 Then
        fraDiagnosa.Visible = False
    End If
End Sub

Private Sub dgDiagnosaKeperawatan_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgDiagnosaKeperawatan
        WheelHook.WheelHook dgDiagnosaKeperawatan
End Sub

Private Sub dgDiagnosaKeperawatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDiagnosa.SetFocus
    If KeyAscii = 27 Then fraDiagnosa.Visible = False
End Sub

Private Sub dgDiagnosaKeperawatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgDiagnosaKeperawatan.ApproxCount = 0 Then Exit Sub
        txtDiagnosa.Text = dgDiagnosaKeperawatan.Columns(1).Value
       txtKdDiagnosaKeperawatan.Text = dgDiagnosaKeperawatan.Columns(0).Value
        dcAskep.BoundText = dgDiagnosaKeperawatan.Columns(2).Value
        txtKodeExternal.Text = dgDiagnosaKeperawatan.Columns(3).Value
        txtNamaExternal.Text = dgDiagnosaKeperawatan.Columns(4).Value
        
        If dgDiagnosaKeperawatan.Columns(5) = "" Then
            Check1.Value = 0
        ElseIf dgDiagnosaKeperawatan.Columns(5) = 0 Then
            Check1.Value = 0
        ElseIf dgDiagnosaKeperawatan.Columns(5) = 1 Then
            Check1.Value = 1
        End If
        fraDiagnosa.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If SSTab1.Tab = 0 Then
    Select Case KeyCode
         Case vbKeyF1
            If dgDiagnosaKeperawatan.ApproxCount = 0 Then Exit Sub
            frmCetakMastDiagnosaKeperawatan.Show
    End Select
   ElseIf SSTab1.Tab = 1 Then
     Select Case KeyCode
        Case vbKeyF1
            If dgDetailAskep.ApproxCount = 0 Then Exit Sub
            frmCetakDetailDiagnosaKeperawatan.Show
    End Select
   
   End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    Call openConnection
    Call clear
    SSTab1.Tab = 0
    Call subLoadDcSource
    Call subLoadGridSource
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadGridSource()
On Error GoTo errLoad
    Set rs = Nothing
 If SSTab1.Tab = 0 Then
   ' strSQL = "select * from DiagnosaKeperawatan"
     strSQL = "Select DiagnosaKeperawatan.*, AsuhanKeperawatan.NamaAskep from DiagnosaKeperawatan inner join AsuhanKeperawatan ON DiagnosaKeperawatan.KdAskep = AsuhanKeperawatan.KdAskep"
     rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgDiagnosaKeperawatan.DataSource = rs
    With dgDiagnosaKeperawatan
        .Columns(0).Caption = "Kd Diagnosa Keperawatan"
        .Columns(0).Width = 0
        .Columns(1).Caption = "Diagnosa Keperawatan"
        .Columns(1).Width = 5900
        .Columns(2).Width = 0
        .Columns(2).Caption = "Kd Askep"
        '.Columns(5).Width = 0
    End With
    Set rs = Nothing
 ElseIf SSTab1.Tab = 1 Then
     Set rs = Nothing
    strSQL = "select * from DetailDiagnosaKeperawatan"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgDetailAskep.DataSource = rs
    With dgDetailAskep
        .Columns(0).Caption = "Kode"
        .Columns(0).Width = 1500
        .Columns(1).Width = 4800
    End With
    Set rs = Nothing
 
 End If
    
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Sub clear()
 If SSTab1.Tab = 0 Then
    txtDiagnosa.Text = ""
   ' txtKdDiagnosaKeperawatan.Text = ""
    dcAskep.Text = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    Check1.Value = 0
  ElseIf SSTab1.Tab = 1 Then
    txtKdDetailAskep.Text = ""
    txtDetailAskep.Text = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    Check1.Value = 1
  
  
  End If
  
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    subLoadGridSource
End Sub

'Private Sub txtDiagnosa_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtKodeExternal.SetFocus
'End Sub

Private Sub txtDiagnosa_LostFocus()
Dim i As Integer
Dim tempText As String

    tempText = Trim(txtDiagnosa.Text)
    txtDiagnosa.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtDiagnosa.Text = txtDiagnosa.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Sub subLoadDiagnosa()
On Error GoTo errLoad
    Set rs = Nothing
    strSQL = "select KdDiagnosa, NamaDiagnosa from Diagnosa " & strFilterDiagnosa
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlDiagnosa = rs.RecordCount
    Set dgDiagnosa.DataSource = rs
    With dgDiagnosa
        .Columns(0).Caption = "Kode Diagnosa"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Nama Diagnosa"
        .Columns(1).Width = 7500
    End With
    fraDiagnosa.Left = 0
    fraDiagnosa.Top = 1920
Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Function sp_DiagnosaKeperawatan(f_Status As String) As Boolean
    sp_DiagnosaKeperawatan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDiagnosaKeperawatan", adVarChar, adParamInput, 5, txtKdDiagnosaKeperawatan.Text)
        .Parameters.Append .CreateParameter("DiagnosaKeperawatan", adVarChar, adParamInput, 1000, Trim(txtDiagnosa.Text))
        .Parameters.Append .CreateParameter("OutputKode", adChar, adParamOutput, 5, Null)
        .Parameters.Append .CreateParameter("KdAskep", adChar, adParamInput, 4, dcAskep.BoundText)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 1000, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , Check1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_DiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_DiagnosaKeperawatan = False
        End If
        
        If f_Status = "A" Then
            txtOutputKode2.Text = .Parameters("OutputKode").Value
            MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
        Else
            MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        End If
        
        Call deleteADOCommandParameters(dbcmd)
        txtOutputKode2.Text = "123"
        Set dbcmd = Nothing
    End With
End Function

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Check1.SetFocus
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub subLoadDcSource()

    Set rs = Nothing
    Call msubDcSource(dcAskep, rs, "Select KdAskep, NamaAskep From AsuhanKeperawatan where StatusEnabled='1' order by NamaAskep")

End Sub
Private Sub dgDetailAskep_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgDetailAskep
        WheelHook.WheelHook dgDetailAskep
End Sub
Private Sub dgDetailAskep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDetailAskep.SetFocus
End Sub


Private Sub dgDetailAskep_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errLoad

    If dgDetailAskep.ApproxCount = 0 Then Exit Sub
    txtKdDetailAskep.Text = dgDetailAskep.Columns(0).Value
    txtDetailAskep.Text = dgDetailAskep.Columns(1).Value
    
    txtKodeExternal.Text = dgDetailAskep.Columns(2).Value
    txtNamaExternal.Text = dgDetailAskep.Columns(3).Value
    If dgDetailAskep.Columns(4) = "" Then
        Check1.Value = 0
    ElseIf dgDetailAskep.Columns(4) = 0 Then
        Check1.Value = 0
    ElseIf dgDetailAskep.Columns(4) = 1 Then
        Check1.Value = 1
    End If
Exit Sub
errLoad:
End Sub


Private Function sp_DetailAskep(f_Status As String) As Boolean
    sp_DetailAskep = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailAskep", adVarChar, adParamInput, 3, txtKdDetailAskep.Text)
        .Parameters.Append .CreateParameter("DetailAskep", adVarChar, adParamInput, 50, Trim(txtDetailAskep.Text))
        .Parameters.Append .CreateParameter("OutputKode", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , Check1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AUD_DetailDiagnosaKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_DetailAskep = False
        Else
            Call Add_HistoryLoginActivity("AUD_DetailDiagnosaKeperawatan")
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
