VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIndikatorPelPerRuang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Indikator Pelayanan Per Ruangan"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   Icon            =   "frmIndikatorPelPerRuang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   7425
   Begin VB.CheckBox chkBulanan 
      Caption         =   "Bulanan"
      Height          =   300
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Periode"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
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
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         CustomFormat    =   "MMMM yyyy"
         Format          =   145948675
         UpDown          =   -1  'True
         CurrentDate     =   38373
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Image Image4 
      Height          =   945
      Left            =   5640
      Picture         =   "frmIndikatorPelPerRuang.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmIndikatorPelPerRuang.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmIndikatorPelPerRuang.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8640
      Picture         =   "frmIndikatorPelPerRuang.frx":5A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmIndikatorPelPerRuang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBulanan_Click()
    If chkBulanan.Value = 1 Then
        dtpAwal.CustomFormat = "MMMM yyyy"
        varStatusBulanan = True
    Else
        dtpAwal.CustomFormat = "yyyy"
        varStatusBulanan = False
    End If
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell

    Dim pesan As VbMsgBoxResult
    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"
    Call MajuTerusPantangMundur(dbcmd)
    frmCtkIndikatorPelPerRuang.Show
hell:

End Sub

Private Sub MajuTerusPantangMundur(ByVal adoCommand As ADODB.Command)
    On Error GoTo hell
    Dim strLokal As String

    If varStatusBulanan = False Then
        varTahun = Format(dtpAwal.Value, "yyyy")
        varBulan = 0
        If varTahun < Year(Now) Then
            varHari = (DateDiff("d", DateValue("01/01/" + varTahun), DateValue("31/12/" + varTahun))) + 1 'lampau
        Else
            varHari = (DateDiff("d", DateValue("01/01/" + varTahun), Now)) + 1 'sekarang
        End If
    Else
        varTahun = Format(dtpAwal.Value, "yyyy")
        varBulan = Format(dtpAwal.Value, "MM")
        If varTahun < Year(Now) Then 'lampau
            varHari = DateDiff("d", DateValue(str(varBulan) + "/" + str(varTahun)), DateValue(str(varBulan + 1) + "/" + str(varTahun)))
        Else
            If varBulan < Month(Now) Then 'lampau
                varHari = DateDiff("d", DateValue(str(varBulan) + "/" + str(varTahun)), DateValue(str(varBulan + 1) + "/" + str(varTahun)))
            Else 'sekarang
                varHari = DateDiff("d", DateValue("01/" + str(varBulan) + "/" + str(varTahun)), Now)
            End If
        End If
    End If
    varTglCetak = Format(dtpAwal.Value, "yyyy/MM/dd 23:59:59")
    If varHari < 1 Then Exit Sub

    Set adoCommand = New ADODB.Command

    MousePointer = vbHourglass
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglHitung", adDate, adParamInput, , Format(varTglCetak, "yyyy/MM/dd 23:59:59"))
        .Parameters.Append .CreateParameter("Hari", adInteger, adParamInput, , varHari)
        .Parameters.Append .CreateParameter("Hasil", adVarChar, adParamOutput, 36, Null)
        .Parameters.Append .CreateParameter("Tahun", adInteger, adParamInput, , varTahun)
        .Parameters.Append .CreateParameter("Bulan", adInteger, adParamInput, , varBulan)
        '

        .ActiveConnection = dbConn
        .CommandText = "Add_ProsesIndikatorX"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Proses Pembuatan Laporan", vbCritical, "Validasi"
        Else

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
        .CommandTimeout = 120000
    End With
    MousePointer = vbDefault
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    dtpAwal.MaxDate = Now
    dtpAwal.Value = Format(Now, "dd MMMM yyyy")
    varStatusBulanan = True
End Sub
