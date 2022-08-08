VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLapIndPlynRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Laporan"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLapIndPlynRS.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14700
   Begin VB.Frame fraPeriode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   14685
      Begin VB.ComboBox cbKriteria 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmLapIndPlynRS.frx":0CCA
         Left            =   480
         List            =   "frmLapIndPlynRS.frx":0CD4
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame Frame1 
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
         Height          =   735
         Left            =   8760
         TabIndex        =   11
         Top             =   150
         Width           =   5775
         Begin VB.CommandButton cmdCari 
            Caption         =   "&Cari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpAwal 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   " MMMM yyyy"
            Format          =   437977091
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   375
            Left            =   3480
            TabIndex        =   1
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   437977091
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "s/d"
            Height          =   210
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.OptionButton opRuangan 
         Caption         =   "Per Ruangan"
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
         Left            =   5160
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton opKelas 
         Caption         =   "Per Kelas"
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
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   4875
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   14445
         _ExtentX        =   25479
         _ExtentY        =   8599
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kriteria Indikator"
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
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   7080
      Width           =   14685
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   11040
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   375
         Left            =   12840
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   14
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
      Left            =   12840
      Picture         =   "frmLapIndPlynRS.frx":0CF0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmLapIndPlynRS.frx":1A78
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmLapIndPlynRS.frx":4439
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmLapIndPlynRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRowNow As Integer
Dim rsTemp1 As ADODB.recordset
Dim rsTemp2 As ADODB.recordset
Dim i As Integer

'' add splakuk 2011/01/24 sp untuk perhitungan indikator pelayanan rs
Private Function sp_IndikatorPelayanan() As Boolean
    sp_IndikatorPelayanan = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("TglHitung", adDate, adParamInput, , Format(dTglHitung, "yyyy/MM/dd 23:59:59"))

        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_IndikatorPelayananRS_New"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_IndikatorPelayanan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

''

Private Sub cmdCari_Click()
    Dim intJmlRow As Integer
    Dim intNo As Integer
    Dim intTgl As Integer

    Set rs = Nothing
    Call msubRecFO(rs, "select * from SKBedRS")
    If rs.EOF Then
        MsgBox "Master SK Bed Rumah Sakit kosong, harap setting dahulu", vbExclamation, "Validasi"
        Exit Sub
    End If

    '' add splakuk 2011/01/24 sp untuk perhitungan indikator pelayanan rs
    '==#######################################
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

    For intTgl = 1 To varHari
        dTglHitung = DateValue(str(intTgl) + "/" + str(varBulan) + "/" + str(varTahun))
        If sp_IndikatorPelayanan() = False Then Exit Sub
    Next intTgl

    Call Add_HistoryLoginActivity("Add_IndikatorPelayananRS_New")

    If cbKriteria.Text = "Per Kelas" Then
        Call subSetGridKelas
    ElseIf cbKriteria.Text = "Per Ruangan" Then
        Call subSetGridRuangan
    End If

    'u/ mempercepat
    fgData.Visible = False
    MousePointer = vbHourglass
    intNo = 0
    iRowNow = 0

    'Hitung jumlah row dari data yang hendak ditampilkan
    If cbKriteria.Text = "Per Kelas" Then
        'jika per kelas
        strSQL = " SELECT TOP 100 PERCENT Kelas, " & _
        " SUM(JmlBed) AS JmlBed, " & _
        " SUM(JmlHariPerawatan) AS JmlHariPerawatan, " & _
        " SUM(JmlPasienOutHidup) AS JmlPasienOutHidup, " & _
        " SUM(JmlPasienOutMati) AS JmlPasienOutMati, " & _
        " SUM(JmlPasienMatiLK48) AS JmlPasienMatiLK48, " & _
        " SUM(JmlPasienMatiLB48) AS JmlPasienMatiLB48, " & _
        " AVG(BOR) AS TBOR, " & _
        " AVG(TOI) AS TTOI, " & _
        " AVG(BTO) AS TBTO, " & _
        " AVG(GDR) AS TGDR, " & _
        " AVG(NDR) AS TNDR " & _
        " From dbo.V_IndikatorPelayananRSPerKelas " & _
        " WHERE (TglHitung BETWEEN ' " & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " & _
        " AND KdRuangan = '" & strNKdRuangan & "' " & _
        " GROUP BY Kelas " & _
        " ORDER BY Kelas "

    ElseIf cbKriteria.Text = "Per Ruangan" Then
        'jika per ruangan
        strSQL = " SELECT TOP 100 PERCENT Ruangan, " & _
        " SUM(JmlBed) AS JmlBed, " & _
        " SUM(JmlHariPerawatan) AS JmlHariPerawatan, " & _
        " SUM(JmlPasienOutHidup) AS JmlPasienOutHidup, " & _
        " SUM(JmlPasienOutMati) AS JmlPasienOutMati, " & _
        " SUM(JmlPasienMatiLK48) AS JmlPasienMatiLK48, " & _
        " SUM(JmlPasienMatiLB48) AS JmlPasienMatiLB48, " & _
        " AVG(LOS) AS TLOS, " & _
        " AVG(BOR) AS TBOR, " & _
        " AVG(TOI) AS TTOI, " & _
        " AVG(BTO) AS TBTO, " & _
        " AVG(GDR) AS TGDR, " & _
        " AVG(NDR) AS TNDR " & _
        " FROM V_IndikatorPelayananRSPerRuangan" & _
        " WHERE (TglHitung BETWEEN ' " & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "') " & _
        " AND KdRuangan = '" & strNKdRuangan & "' " & _
        " GROUP BY Ruangan" & _
        " ORDER BY Ruangan"
    End If

    msubRecFO rs, strSQL
    'jika tidak ada data
    If rs.EOF = True Then
        fgData.Visible = True
        MousePointer = vbNormal
        MsgBox "Tidak ada Data", vbExclamation, "Validasi"
        Exit Sub
    End If
    intJmlRow = rs.RecordCount + 2

    'u/ menampilkan yang di group by
    With fgData
        'jml baris akhir
        .Rows = intJmlRow
        While rs.EOF = False
            'baris u/ sub total
            iRowNow = iRowNow + 1
            intNo = intNo + 1
            .TextMatrix(iRowNow, 0) = intNo

            If cbKriteria.Text = "Per Kelas" Then
                .TextMatrix(iRowNow, 1) = IIf(rs("Kelas").Value = 0, "0,00", Format(rs("Kelas").Value, "#,###.00"))
                .TextMatrix(iRowNow, 2) = IIf(rs("JmlBed").Value = 0, "0,00", Format(rs("JmlBed").Value, "#,###.00"))
                .TextMatrix(iRowNow, 3) = IIf(rs("JmlHariPerawatan").Value = 0, "0,00", Format(rs("JmlHariPerawatan").Value, "#,###.00"))
                .TextMatrix(iRowNow, 4) = IIf(rs("JmlPasienOutHidup").Value = 0, "0,00", Format(rs("JmlPasienOutHidup").Value, "#,###.00"))
                .TextMatrix(iRowNow, 5) = IIf(rs("JmlPasienOutMati").Value = 0, "0,00", Format(rs("JmlPasienOutMati").Value, "#,###.00"))
                .TextMatrix(iRowNow, 6) = IIf(rs("JmlPasienMatiLK48").Value = 0, "0,00", Format(rs("JmlPasienMatiLK48").Value, "#,###.00"))
                .TextMatrix(iRowNow, 7) = IIf(rs("JmlPasienMatiLB48").Value = 0, "0,00", Format(rs("JmlPasienMatiLB48").Value, "#,###.00"))
                .TextMatrix(iRowNow, 8) = IIf(rs("TBOR").Value = 0, "0,00", Format(rs("TBOR").Value, "#,###.00"))
                .TextMatrix(iRowNow, 9) = IIf(rs("TTOI").Value = 0, "0,00", Format(rs("TTOI").Value, "#,###.00"))
                .TextMatrix(iRowNow, 10) = IIf(rs("TBTO").Value = 0, "0,00", Format(rs("TBTO").Value, "#,###.00"))
                .TextMatrix(iRowNow, 11) = IIf(rs("TGDR").Value = 0, "0,00", Format(rs("TGDR").Value, "#,###.00"))
                .TextMatrix(iRowNow, 12) = IIf(rs("TNDR").Value = 0, "0,00", Format(rs("TNDR").Value, "#,###.00"))
            ElseIf cbKriteria.Text = "Per Ruangan" Then
                .TextMatrix(iRowNow, 1) = rs("Ruangan").Value
                .TextMatrix(iRowNow, 2) = IIf(rs("JmlBed").Value = 0, "0,00", Format(rs("JmlBed").Value, "#,###.00"))
                .TextMatrix(iRowNow, 3) = IIf(rs("JmlHariPerawatan").Value = 0, "0,00", Format(rs("JmlHariPerawatan").Value, "#,###.00"))
                .TextMatrix(iRowNow, 4) = IIf(rs("JmlPasienOutHidup").Value = 0, "0,00", Format(rs("JmlPasienOutHidup").Value, "#,###.00"))
                .TextMatrix(iRowNow, 5) = IIf(rs("JmlPasienOutMati").Value = 0, "0,00", Format(rs("JmlPasienOutMati").Value, "#,###.00"))
                .TextMatrix(iRowNow, 6) = IIf(rs("JmlPasienMatiLK48").Value = 0, "0,00", Format(rs("JmlPasienMatiLK48").Value, "#,###.00"))
                .TextMatrix(iRowNow, 7) = IIf(rs("JmlPasienMatiLB48").Value = 0, "0,00", Format(rs("JmlPasienMatiLB48").Value, "#,###.00"))
                .TextMatrix(iRowNow, 8) = IIf(rs("TLOS").Value = 0, "0,00", Format(rs("TLOS").Value, "#,###.00"))
                .TextMatrix(iRowNow, 9) = IIf(rs("TBOR").Value = 0, "0,00", Format(rs("TBOR").Value, "#,###.00"))
                .TextMatrix(iRowNow, 10) = IIf(rs("TTOI").Value = 0, "0,00", Format(rs("TTOI").Value, "#,###.00"))
                .TextMatrix(iRowNow, 11) = IIf(rs("TBTO").Value = 0, "0,00", Format(rs("TBTO").Value, "#,###.00"))
                .TextMatrix(iRowNow, 12) = IIf(rs("TGDR").Value = 0, "0,00", Format(rs("TGDR").Value, "#,###.00"))
                .TextMatrix(iRowNow, 13) = IIf(rs("TNDR").Value = 0, "0,00", Format(rs("TNDR").Value, "#,###.00"))
            End If
            rs.MoveNext
        Wend
        iRowNow = iRowNow + 1
        If cbKriteria.Text = "Per Kelas" Then
            strSQL = "SELECT SUM(JmlBed) AS TJmlBed, " & _
            " SUM(JmlHariPerawatan) AS TJmlHariPerawatan, " & _
            " SUM(JmlPasienOutHidup) AS TJmlPasienOutHidup, " & _
            " SUM(JmlPasienOutMati) AS TJmlPasienOutMati, " & _
            " SUM(JmlPasienMatiLK48) AS TJmlPasienMatiLK48, " & _
            " SUM(JmlPasienMatiLB48) AS TJmlPasienMatiLB48, " & _
            " AVG(BOR) AS TBOR, " & _
            " AVG(TOI) AS TTOI, " & _
            " AVG(BTO) AS TBTO, " & _
            " AVG(GDR) AS TGDR, " & _
            " AVG(NDR) AS TNDR " & _
            " FROM V_IndikatorPelayananRSPerKelas" & _
            " WHERE (TglHitung BETWEEN ' " & Format(dtpAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir.Value, "yyyy/MM/dd 23:59:59") & "')" & _
            " AND KdRuangan = '" & strNKdRuangan & "' "
            msubOpenRecFO rs, strSQL, dbConn
            .TextMatrix(iRowNow, 1) = "Total/Rata2"
            .TextMatrix(iRowNow, 2) = IIf(rs("TJmlBed").Value = 0, "0,00", Format(rs("TJmlBed").Value, "#,###.00"))
            .TextMatrix(iRowNow, 3) = IIf(rs("TJmlHariPerawatan").Value = 0, "0,00", Format(rs("TJmlHariPerawatan").Value, "#,###.00"))
            .TextMatrix(iRowNow, 4) = IIf(rs("TJmlPasienOutHidup").Value = 0, "0,00", Format(rs("TJmlPasienOutHidup").Value, "#,###.00"))
            .TextMatrix(iRowNow, 5) = IIf(rs("TJmlPasienOutMati").Value = 0, "0,00", Format(rs("TJmlPasienOutMati").Value, "#,###.00"))
            .TextMatrix(iRowNow, 6) = IIf(rs("TJmlPasienMatiLK48").Value = 0, "0,00", Format(rs("TJmlPasienMatiLK48").Value, "#,###.00"))
            .TextMatrix(iRowNow, 7) = IIf(rs("TJmlPasienMatiLB48").Value = 0, "0,00", Format(rs("TJmlPasienMatiLB48").Value, "#,###.00"))
            .TextMatrix(iRowNow, 8) = IIf(rs("TBOR").Value = 0, "0,00", Format(rs("TBOR").Value, "#,###.00"))
            .TextMatrix(iRowNow, 9) = IIf(rs("TTOI").Value = 0, "0,00", Format(rs("TTOI").Value, "#,###.00"))
            .TextMatrix(iRowNow, 10) = IIf(rs("TBTO").Value = 0, "0,00", Format(rs("TBTO").Value, "#,###.00"))
            .TextMatrix(iRowNow, 11) = IIf(rs("TGDR").Value = 0, "0,00", Format(rs("TGDR").Value, "#,###.00"))
            .TextMatrix(iRowNow, 12) = IIf(rs("TNDR").Value = 0, "0,00", Format(rs("TNDR").Value, "#,###.00"))

            subSetSubTotalRow iRowNow, 1, vbBlue, vbWhite

        End If
    End With
    fgData.Visible = True
    MousePointer = vbNormal

    cmdCetak.SetFocus
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    'On Error Resume Next
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    mblnGrafik = False

    If cbKriteria.Text = "Per Kelas" Then
        strSQL = "SELECT Ruangan FROM V_IndikatorPelayananRSPerKelas " _
        & "WHERE (TglHitung BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & "AND KdRuangan='" & strNKdRuangan & "'"
    ElseIf cbKriteria.Text = "Per Ruangan" Then
        strSQL = "SELECT Ruangan FROM V_IndikatorPelayananRSPerRuangan " _
        & "WHERE (TglHitung BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & "AND KdRuangan='" & strNKdRuangan & "'"
    End If
    msubRecFO rs, strSQL

    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    Else
        vLaporan = ""
        If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
        frmCetakLaporanIndPlynRS.Show
        If cbKriteria.Text = "Per Kelas" = True Then
            frmCetakLaporanIndPlynRS.Caption = "Medifirst2000 - Indikator Pelayanan RS Per Kelas"
        Else
            frmCetakLaporanIndPlynRS.Caption = "Medifirst2000 - Indikator Pelayanan RS Per Ruangan"
        End If
        cmdCetak.Enabled = True
    End If
hell:
    '    msubPesanError
End Sub

Private Sub cmdgrafik_Click()
    cmdCetak.Enabled = False
    mdTglAwal = dtpAwal.Value
    mdTglAkhir = dtpAkhir.Value
    mblnGrafik = True

    If cbKriteria.Text = "Per Kelas" = True Then
        strSQL = "SELECT Ruangan FROM V_IndikatorPelayananRSPerKelas " _
        & "WHERE (TglHitung BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & "AND KdRuangan='" & mstrKdRuangan & "'"
    ElseIf cbKriteria.Text = "Per Ruangan" = True Then
        strSQL = "SELECT Ruangan FROM V_IndikatorPelayananRSPerRuangan " _
        & "WHERE (TglHitung BETWEEN '" _
        & Format(mdTglAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(mdTglAkhir, "yyyy/MM/dd 23:59:59") & "') " _
        & "AND KdRuangan='" & mstrKdRuangan & "'"
    End If
    msubRecFO rs, strSQL

    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If

    frmCetakLaporanIndPlynRS.Show
    If cbKriteria.Text = "Per Kelas" = True Then
        frmCetakLaporanIndPlynRS.Caption = "Medifirst2000 - Indikator Pelayanan RS Per Kelas"
    Else
        frmCetakLaporanIndPlynRS.Caption = "Medifirst2000 - Indikator Pelayanan RS Per Ruangan"
    End If
    cmdCetak.Enabled = True
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dtpAkhir_Change()
    dtpAkhir.MaxDate = Now
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then opKelas.SetFocus
End Sub

Private Sub dtpAwal_Change()
    dtpAwal.MaxDate = Now
End Sub

Private Sub dtpAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    With Me
        .dtpAwal.Value = Now
        .dtpAkhir.Value = Now
        .cbKriteria.Text = "Per Kelas"
    End With
    Call subSetGridKelas
End Sub

'Untuk setting grid per ruangan
Private Sub subSetGridRuangan()
    With fgData
        .Visible = False
        .clear
        .Cols = 14
        .Rows = 2
        .Row = 0

        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .MergeCells = 1
        .MergeCol(1) = True

        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Ruangan"
        .TextMatrix(0, 2) = "JmlBed"
        .TextMatrix(0, 3) = "JmlHariPerawatan"
        .TextMatrix(0, 4) = "JmlPasienOutHidup"
        .TextMatrix(0, 5) = "JmlPasienOutMati"
        .TextMatrix(0, 6) = "JmlPasienMati<48"
        .TextMatrix(0, 7) = "JmlPasienMati>48"
        .TextMatrix(0, 8) = "LOS"
        .TextMatrix(0, 9) = "BOR"
        .TextMatrix(0, 10) = "TOI"
        .TextMatrix(0, 11) = "BTO"
        .TextMatrix(0, 12) = "GDR"
        .TextMatrix(0, 13) = "NDR"

        .ColWidth(0) = 500
        .ColWidth(1) = 2600
        .ColWidth(2) = 1100
        .ColWidth(3) = 1800
        .ColWidth(4) = 1900
        .ColWidth(5) = 1800
        .ColWidth(6) = 1800
        .ColWidth(7) = 1800
        .ColWidth(8) = 1100
        .ColWidth(9) = 1100
        .ColWidth(10) = 1100
        .ColWidth(11) = 1100
        .ColWidth(12) = 1100
        .ColWidth(13) = 1100

        .Visible = True
        iRowNow = 0
    End With
End Sub

'Untuk setting grid per kelas
Private Sub subSetGridKelas()
    With fgData
        .Visible = False
        .clear
        .Cols = 13
        .Rows = 2
        .Row = 0

        For i = 0 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next

        .MergeCells = 1
        .MergeCol(1) = True

        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Kelas"
        .TextMatrix(0, 2) = "JmlBed"
        .TextMatrix(0, 3) = "JmlHariPerawatan"
        .TextMatrix(0, 4) = "JmlPasienOutHidup"
        .TextMatrix(0, 5) = "JmlPasienOutMati"
        .TextMatrix(0, 6) = "JmlPasienMati<48"
        .TextMatrix(0, 7) = "JmlPasienMati>48"
        .TextMatrix(0, 8) = "BOR"
        .TextMatrix(0, 9) = "TOI"
        .TextMatrix(0, 10) = "BTO"
        .TextMatrix(0, 11) = "GDR"
        .TextMatrix(0, 12) = "NDR"

        .ColWidth(0) = 500
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1800
        .ColWidth(4) = 1900
        .ColWidth(5) = 1800
        .ColWidth(6) = 1800
        .ColWidth(7) = 1800
        .ColWidth(8) = 1100
        .ColWidth(9) = 1100
        .ColWidth(10) = 1100
        .ColWidth(11) = 1100
        .ColWidth(12) = 1100

        .Visible = True
        iRowNow = 0
    End With
End Sub

'Untuk mensetting grid di row subtotal
Private Sub subSetSubTotalRow(iRowNow As Integer, iColMulai As Integer, vbBackColor, vbForeColor)
    Dim i As Integer
    With fgData
        'tampilan Black & White
        For i = iColMulai To .Cols - 1
            .Col = i
            .Row = iRowNow
            .CellBackColor = vbBackColor
            .CellForeColor = vbForeColor
            .CellFontBold = True
        Next
    End With
End Sub

Private Sub opKelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

Private Sub opRuangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdCari.SetFocus
End Sub

