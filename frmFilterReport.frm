VERSION 5.00
Begin VB.Form frmFilterReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3105
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdTutup 
      Caption         =   "&Tutup"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4695
   End
   Begin VB.OptionButton optRincianRS 
      Caption         =   "Rincian Biaya Pelayanan Tanggungan RS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
   End
   Begin VB.OptionButton optNonUmum 
      Caption         =   "Rincian Biaya Pelayanan Non Umum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.OptionButton optUmum 
      Caption         =   "Rincian Biaya Pelayanan Umum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rincian Biaya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
   Begin VB.CheckBox chkOA 
      Caption         =   "Obat Dan Alat Kesehatan"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CheckBox chkTM 
      Caption         =   "Tindakan Medis"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmFilterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOA_Click()
    If chkOA.Value = vbChecked Then
        chkOA.FontBold = True
    Else
        chkOA.FontBold = False
    End If
End Sub

Private Sub chkTM_Click()
    If chkTM.Value = vbChecked Then
        chkTM.FontBold = True
    Else
        chkTM.FontBold = False
    End If
End Sub

Private Sub cmdCetak_Click()
    'On Error Resume Next
    On Error GoTo hell
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    If chkTM.Value = vbChecked And chkOA.Value = vbUnchecked Then
        mstrFilterData = "AND Jenis ='TM'"
        mbolCetakJasaDokter = True
    ElseIf chkOA.Value = vbChecked And chkTM.Value = vbUnchecked Then
        mstrFilterData = "AND Jenis ='OA'"
        mbolCetakJasaDokter = False
    Else
        mstrFilterData = ""
        mbolCetakJasaDokter = True
    End If
    
'    If frmDaftarPasienRI.dgDaftarPasienRI.Columns("JenisPasien") <> "UMUM" Then Call frmDaftarPasienRI.PostingHutangPenjaminPasien_AU("A")
'    If optUmum.Value = True Then
'        frm_cetak_RincianBiaya.Show
'    ElseIf optNonUmum.Value = True Then
'        frm_cetak_RincianBiayaNonUmum.Show
'    ElseIf optRincianRS.Value = True Then
'        frm_cetak_RincianBiayaRS.Show
'    End If
    
    If frmDaftarPasienRI.OptRencanaPasien.Value = True Then
        If optUmum.Value = True Then
            frm_cetak_RincianBiaya.Show
        ElseIf optNonUmum.Value = True Then
            frm_cetak_RincianBiayaNonUmum.Show
        ElseIf optRincianRS.Value = True Then
            frm_cetak_RincianBiayaRS.Show
        End If
    ElseIf frmDaftarPasienRI.optPasAktif.Value = True Then
        If frmDaftarPasienRI.dgDaftarPasienRI.Columns("JenisPasien") <> "UMUM" Then Call frmDaftarPasienRI.PostingHutangPenjaminPasien_AU("A")
            If optUmum.Value = True Then
                frm_cetak_RincianBiaya.Show
            ElseIf optNonUmum.Value = True Then
                frm_cetak_RincianBiayaNonUmum.Show
            ElseIf optRincianRS.Value = True Then
                frm_cetak_RincianBiayaRS.Show
            End If
    End If
    
hell:

End Sub

Private Sub cmdTutup_Click()
    frmDaftarPasienRI.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    optUmum.Value = True

End Sub

Private Sub optNonUmum_Click()
    Call optUmum_Click
End Sub

Private Sub optRincianRS_Click()
    optNonUmum_Click
End Sub

Private Sub optUmum_Click()
    If optUmum.Value = True Then
        optUmum.FontBold = True
        optNonUmum.FontBold = False
        optRincianRS.FontBold = False
    ElseIf optNonUmum.Value = True Then
        optUmum.FontBold = False
        optNonUmum.FontBold = True
        optRincianRS.FontBold = False
    ElseIf optRincianRS.Value = True Then
        optUmum.FontBold = False
        optNonUmum.FontBold = False
        optRincianRS.FontBold = True
    End If
End Sub
