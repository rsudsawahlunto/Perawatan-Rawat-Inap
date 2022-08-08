VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterAsuhanKeperawatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Asuhan Keperawatan"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterAsuhanKeperawatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   10635
   Begin VB.TextBox txtOuputKode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
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
      Top             =   7680
      Width           =   10575
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   4830
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
   Begin VB.TextBox txtKdDiagnosaKeperawatan 
      Height          =   375
      Left            =   0
      MaxLength       =   10
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
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
      Height          =   6375
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "Kategory Asuhan Keperawatan"
      TabPicture(0)   =   "frmMasterAsuhanKeperawatan.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Asuhan Keperawatan"
      TabPicture(1)   =   "frmMasterAsuhanKeperawatan.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
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
         Height          =   5775
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   9975
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5295
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   9735
            Begin VB.TextBox txtKodeExternal 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   240
               TabIndex        =   17
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtNamaExternal 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2280
               TabIndex        =   16
               Top             =   1320
               Width           =   5535
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   7920
               TabIndex        =   15
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtKategoryAskep 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   240
               MaxLength       =   50
               TabIndex        =   14
               Top             =   600
               Width           =   6135
            End
            Begin VB.TextBox txtKdKategoryAskep 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   240
               MaxLength       =   3
               TabIndex        =   13
               Top             =   600
               Width           =   735
            End
            Begin MSDataGridLib.DataGrid dgKategoryAskep 
               Height          =   3135
               Left            =   120
               TabIndex        =   18
               Top             =   1680
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   5530
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
            Begin VB.Label Label6 
               Caption         =   "Kode External"
               Height          =   255
               Left            =   240
               TabIndex        =   22
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Nama External"
               Height          =   255
               Left            =   2280
               TabIndex        =   21
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Kategory Asuhan Keperawatan"
               Height          =   210
               Left            =   240
               TabIndex        =   20
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Kode"
               Height          =   210
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   420
            End
         End
      End
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
         Height          =   5895
         Left            =   -74760
         TabIndex        =   10
         Top             =   360
         Width           =   10215
         Begin VB.Frame Frame4 
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
            TabIndex        =   23
            Top             =   240
            Width           =   9855
            Begin VB.TextBox txtDeskripsi 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               TabIndex        =   40
               Top             =   1320
               Width           =   2775
            End
            Begin VB.TextBox txtKomplikasi 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               TabIndex        =   39
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox txtTandaGejala 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   240
               TabIndex        =   37
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox txtPenyebab 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6240
               TabIndex        =   33
               Top             =   1320
               Width           =   2775
            End
            Begin VB.TextBox txtKdAsuhanKeperawatan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   8520
               MaxLength       =   4
               TabIndex        =   27
               Text            =   "1234"
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtAsuhanKeperawatan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   240
               MaxLength       =   50
               TabIndex        =   26
               Top             =   600
               Width           =   6135
            End
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   6840
               TabIndex        =   25
               Top             =   600
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtPenatalaksanaan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6240
               TabIndex        =   24
               Top             =   2040
               Width           =   2775
            End
            Begin MSDataGridLib.DataGrid dgAsuhanKeperawatan 
               Height          =   2895
               Left            =   240
               TabIndex        =   28
               Top             =   2520
               Width           =   9480
               _ExtentX        =   16722
               _ExtentY        =   5106
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
            Begin MSDataListLib.DataCombo dcKategoryAskep 
               Height          =   330
               Left            =   240
               TabIndex        =   31
               Top             =   1320
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label Label14 
               Caption         =   "Komplikasi"
               Height          =   255
               Left            =   3240
               TabIndex        =   38
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Tanda Gejala"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Deskripsi"
               Height          =   255
               Left            =   3240
               TabIndex        =   35
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Penyebab"
               Height          =   255
               Left            =   6240
               TabIndex        =   34
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label10 
               Caption         =   "Kategory Askep"
               Height          =   255
               Left            =   240
               TabIndex        =   32
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Asuhan Keperawatan"
               Height          =   210
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label2 
               Caption         =   "Penatalaksanaam"
               Height          =   255
               Left            =   6240
               TabIndex        =   29
               Top             =   1800
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8760
      Picture         =   "frmMasterAsuhanKeperawatan.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterAsuhanKeperawatan.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterAsuhanKeperawatan.frx":30E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterAsuhanKeperawatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilterDiagnosa As String
Dim intJmlDiagnosa As Integer


Private Sub cmdBatal_Click()
    Call clear
    Call subLoadGridSource
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errLoad
 If SSTab1.Tab = 0 Then

    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_KategoryAsuhanKeperawatan("D") = False Then Exit Sub
    Call cmdBatal_Click
 
 ElseIf SSTab1.Tab = 1 Then
    On Error GoTo errLoad
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If sp_AsuhanKeperawatan("D") = False Then Exit Sub
    Call cmdBatal_Click
    
 End If
 
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
  If SSTab1.Tab = 0 Then
    If Periksa("text", txtKategoryAskep, "Kategory Asuhan Keperawatan kosong") = False Then Exit Sub
    If sp_KategoryAsuhanKeperawatan("A") = False Then Exit Sub
    Call cmdBatal_Click
  ElseIf SSTab1.Tab = 1 Then
    On Error GoTo errLoad
    If Periksa("text", txtAsuhanKeperawatan, "Asuhan Keperawatan kosong") = False Then Exit Sub
    If Periksa("datacombo", dcKategoryAskep, "Kategory Asuhan Keperawatan kosong") = False Then Exit Sub
    If sp_AsuhanKeperawatan("A") = False Then Exit Sub
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

'Private Sub dgKategoryAskep_Click()
'  WheelHook.WheelUnHook
'  Set MyProperty = dgDiagnosa
'  WheelHook.WheelHook dgDiagnosa
'End Sub

Private Sub dgKategoryAskep_DblClick()
    Call dgKategoryAskep_KeyPress(13)
End Sub
'
'Private Sub dgDiagnosaKeperawatan_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgDiagnosaKeperawatan
'        WheelHook.WheelHook dgDiagnosaKeperawatan
'End Sub


Private Sub dgKategoryAskep_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgKategoryAskep.ApproxCount = 0 Then Exit Sub
        txtKategoryAskep.Text = dgKategoryAskep.Columns(1).Value
        txtKdKategoryAskep.Text = dgKategoryAskep.Columns(0).Value
        txtKodeExternal.Text = dgKategoryAskep.Columns(2).Value
        txtNamaExternal.Text = dgKategoryAskep.Columns(3).Value
        If dgKategoryAskep.Columns(4) = "" Then
            Check1.Value = 0
        ElseIf dgKategoryAskep.Columns(4) = 0 Then
            Check1.Value = 0
        ElseIf dgKategoryAskep.Columns(4) = 1 Then
            Check1.Value = 1
        End If
End Sub


Private Sub dgAsuhanKeperawatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgAsuhanKeperawatan.ApproxCount = 0 Then Exit Sub
        txtKdAsuhanKeperawatan.Text = dgAsuhanKeperawatan.Columns(0).Value
        txtAsuhanKeperawatan.Text = dgAsuhanKeperawatan.Columns(1).Value
        dcKategoryAskep.BoundText = dgAsuhanKeperawatan.Columns(2).Value
        txtDeskripsi.Text = dgAsuhanKeperawatan.Columns(3).Value
        txtPenyebab.Text = dgAsuhanKeperawatan.Columns(4).Value
        txtTandaGejala.Text = dgAsuhanKeperawatan.Columns(5).Value
        txtKomplikasi.Text = dgAsuhanKeperawatan.Columns(6).Value
        txtPenatalaksanaan.Text = dgAsuhanKeperawatan.Columns(7).Value
        
        If dgAsuhanKeperawatan.Columns(8) = "" Then
            Check2.Value = 0
        ElseIf dgAsuhanKeperawatan.Columns(8) = 0 Then
            Check2.Value = 0
        ElseIf dgAsuhanKeperawatan.Columns(8) = 1 Then
            Check2.Value = 1
        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If SSTab1.Tab = 0 Then
    Select Case KeyCode
         Case vbKeyF1
            If dgKategoryAskep.ApproxCount = 0 Then Exit Sub
'            frmCetakMastDiagnosaKeperawatan.Show
    End Select
   ElseIf SSTab1.Tab = 1 Then
     Select Case KeyCode
        Case vbKeyF1
            If dgAsuhanKeperawatan.ApproxCount = 0 Then Exit Sub
'            frmCetakDetailDiagnosaKeperawatan.Show
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

Private Sub subLoadDcSource()

    Set rs = Nothing
    Call msubDcSource(dcKategoryAskep, rs, "Select KdKategoryAskep, KategoryAskep From KategoryAsuhanKeperawatan where StatusEnabled='1' order by KategoryAskep")

End Sub
Private Sub subLoadGridSource()
On Error GoTo errLoad
    Set rs = Nothing
 If SSTab1.Tab = 0 Then
     strSQL = "Select * from KategoryAsuhanKeperawatan"
     rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgKategoryAskep.DataSource = rs
    With dgKategoryAskep
        .Columns(0).Caption = "KdKategoryAskep"
        .Columns(0).Width = 0
        .Columns(1).Caption = "Kategory Askep"
        .Columns(1).Width = 3000
        .Columns(2).Width = 0
    End With
    Set rs = Nothing
 ElseIf SSTab1.Tab = 1 Then
'    Set rs = Nothing
    strSQL = "select * from AsuhanKeperawatan"
    rs.Open strSQL, dbConn, adOpenDynamic, adLockOptimistic
    Set dgAsuhanKeperawatan.DataSource = rs
    With dgAsuhanKeperawatan
        .Columns(0).Caption = "KdAskep"
        .Columns(0).Width = 0
        .Columns(1).Width = 2000
        .Columns(1).Caption = "Nama Askep"
        .Columns(2).Width = 0
        .Columns(2).Caption = "KdKategoryAskep"
        .Columns(3).Width = 2000
        .Columns(3).Caption = "Deskripsi"
        .Columns(4).Width = 2000
        .Columns(4).Caption = "Penyebab"
        .Columns(5).Width = 2000
        .Columns(5).Caption = "Tanda Gejala"
        .Columns(6).Width = 2000
        .Columns(6).Caption = "Komplikasi"
        .Columns(7).Width = 2000
        .Columns(7).Caption = "Penatalaksanaan"
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
    txtKategoryAskep.Text = ""
    txtKodeExternal.Text = ""
    txtNamaExternal.Text = ""
    Check1.Value = 0
  ElseIf SSTab1.Tab = 1 Then
   txtKdAsuhanKeperawatan.Text = ""
    txtAsuhanKeperawatan.Text = ""
    dcKategoryAskep.Text = ""
    txtDeskripsi.Text = ""
    txtPenyebab.Text = ""
    txtTandaGejala.Text = ""
    txtKomplikasi.Text = ""
    txtPenatalaksanaan.Text = ""
    Check2.Value = 0
  End If
  
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    subLoadGridSource
     Call clear
     Call subLoadDcSource
     Check1.Value = 1
     Check2.Value = 1
End Sub

Private Sub txtKategoryAskep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKodeExternal.SetFocus
End Sub

Private Sub txtKategoryAskep_LostFocus()
Dim i As Integer
Dim tempText As String

    tempText = Trim(txtKategoryAskep.Text)
    txtKategoryAskep.Text = ""
    For i = 1 To Len(tempText)
        If Asc(Mid(tempText, i, 1)) <> 10 And Asc(Mid(tempText, i, 1)) <> 13 Then
            txtKategoryAskep.Text = txtKategoryAskep.Text & Mid(tempText, i, 1)
        End If
    Next i
End Sub

Private Function sp_KategoryAsuhanKeperawatan(f_Status As String) As Boolean
    sp_KategoryAsuhanKeperawatan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKategoryAskep", adTinyInt, adParamInput, , IIf(txtKdKategoryAskep.Text = "", Null, txtKdKategoryAskep.Text))
        .Parameters.Append .CreateParameter("KategoryAskep", adVarChar, adParamInput, 20, Trim(txtKategoryAskep.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKodeExternal.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExternal.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , Check1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_KategoryAsuhanKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("return_value").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data, hubungi administrator", vbCritical
            sp_KategoryAsuhanKeperawatan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function


Private Function sp_AsuhanKeperawatan(f_Status As String) As Boolean
    sp_AsuhanKeperawatan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
       ' .Parameters.Append .CreateParameter("KdAskep", adChar, adParamInput, 4, IIf(txtKdAsuhanKeperawatan.Text = "", Null, txtKdAsuhanKeperawatan.Text))
        .Parameters.Append .CreateParameter("KdAskep", adChar, adParamInput, 4, txtKdAsuhanKeperawatan.Text)
        .Parameters.Append .CreateParameter("NamaAskep", adVarChar, adParamInput, 50, Trim(txtAsuhanKeperawatan.Text))
        .Parameters.Append .CreateParameter("KdKategoryAskep", adTinyInt, adParamInput, , dcKategoryAskep.BoundText)
        .Parameters.Append .CreateParameter("Deskripsi", adVarChar, adParamInput, 150, txtDeskripsi.Text)
        .Parameters.Append .CreateParameter("Penyebab", adVarChar, adParamInput, 150, txtPenyebab.Text)
        .Parameters.Append .CreateParameter("TandaGejala", adVarChar, adParamInput, 150, txtTandaGejala.Text)
        .Parameters.Append .CreateParameter("Komplikasi", adVarChar, adParamInput, 150, txtKomplikasi.Text)
        .Parameters.Append .CreateParameter("Penatalaksanaan", adVarChar, adParamInput, 150, txtPenatalaksanaan.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , Check2.Value)
        .Parameters.Append .CreateParameter("OutputKode", adChar, adParamOutput, 4, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_AsuhanKeperawatan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            If f_Status = "A" Then
                MsgBox "Gagal menyimpan data", vbCritical, "Validasi"
            Else
                MsgBox "Gagal menghapus data", vbCritical, "Validasi"
            End If
            sp_AsuhanKeperawatan = False
        Else
            Call Add_HistoryLoginActivity("AUD_AsuhanKeperawatan")
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

Private Sub txtKodeExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExternal.SetFocus
End Sub

Private Sub txtNamaExternal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Check1.SetFocus
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

'
'Private Sub dgDetailAskep_Click()
'WheelHook.WheelUnHook
'        Set MyProperty = dgDetailAskep
'        WheelHook.WheelHook dgDetailAskep
'End Sub
Private Sub dgKategoryAskep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKategoryAskep.SetFocus
End Sub

