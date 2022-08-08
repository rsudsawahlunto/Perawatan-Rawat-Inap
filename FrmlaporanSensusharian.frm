VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmlaporanSensusharian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Lapoan Sensus Harian"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmlaporanSensusharian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   15345
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   15135
      Begin VB.CommandButton cmdcari 
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
         Left            =   8400
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpAwal 
         Height          =   375
         Left            =   9360
         TabIndex        =   9
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
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
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   125960195
         UpDown          =   -1  'True
         CurrentDate     =   37956
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   375
         Left            =   12240
         TabIndex        =   10
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
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
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   125960195
         UpDown          =   -1  'True
         CurrentDate     =   37956
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Ruangan Pelayanan"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
         Height          =   210
         Left            =   11760
         TabIndex        =   11
         Top             =   435
         Width           =   255
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   15135
      Begin VB.OptionButton optResume 
         Caption         =   "Resume Sensus"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         Height          =   495
         Left            =   13320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "Ceta&k"
         Height          =   495
         Left            =   11520
         TabIndex        =   3
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   15135
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   5475
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14925
         _ExtentX        =   26326
         _ExtentY        =   9657
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
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
      Left            =   13560
      Picture         =   "FrmlaporanSensusharian.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "FrmlaporanSensusharian.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "FrmlaporanSensusharian.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "FrmlaporanSensusharian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub cmdCari_Click()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim o As Integer
    Dim p As Integer
    Dim q As Integer

    Dim ipm As Integer
    Dim ipp As Integer
    Dim ipd As Integer
    Dim ipk As Integer

    If dcRuangan.Text <> "" Then
        Call subSetGridClear

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas From LaporanSensusHarianRIPasienMasuk_V " & _
        "where  TglSensus BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & dcRuangan.BoundText & "%' and kdcaramasuk <>'04' ORDER BY RuanganPelayanan"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = False Then

            ipm = rs.RecordCount

            For i = 1 To rs.RecordCount
                With fgData
                    .Rows = rs.RecordCount + 1
                    .TextMatrix(i, 0) = IIf(IsNull(rs.Fields(0).Value), "-", rs.Fields(0))
                    .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(1).Value), "-", rs.Fields(1))
                    .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(2).Value), "-", rs.Fields(2))
                    .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(3).Value), "-", rs.Fields(3))
                    .TextMatrix(i, 28) = 1

                End With
                rs.MoveNext
            Next i
        End If
        'Pasien Pindahan

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas,NoPendaftaran,NoPakai,KdRuanganTujuan,RuanganTujuan From LaporanSensusHarianRIPasienDiPindahkan_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganTujuan Like '%" & dcRuangan.BoundText & "%' ORDER BY RuanganTujuan"
        '
        Set rsB = Nothing
        rsB.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        If rsB.EOF = False Then
            ipp = rsB.RecordCount

            If ipm >= ipp Then
                For j = 1 To rsB.RecordCount
                    With fgData
                        .Rows = rs.RecordCount + 1
                        .TextMatrix(j, 4) = IIf(IsNull(rsB.Fields(0).Value), "-", rsB.Fields(0))
                        .TextMatrix(j, 5) = IIf(IsNull(rsB.Fields(1).Value), "-", rsB.Fields(1))
                        .TextMatrix(j, 6) = IIf(IsNull(rsB.Fields(2).Value), "-", rsB.Fields(2))
                        .TextMatrix(j, 7) = IIf(IsNull(rsB.Fields(3).Value), "-", rsB.Fields(3))
                        .TextMatrix(j, 8) = IIf(IsNull(rsB.Fields(7).Value), "-", rsB.Fields(7)) 'RuanganTujuanM
                        strSQLy = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsB("NoPendaftaran") & "' and KdRuangan = '" & rsB.Fields("KdRuanganTujuan") & "'"
                        Set rsE = Nothing
                        rsE.Open strSQLy, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsE.EOF = True Then
                            .TextMatrix(j, 9) = "Belum Masuk Ruangan"
                            .TextMatrix(j, 28) = 1

                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsE("KdKelasPel") & "'"
                            Set rsE = Nothing
                            rsE.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(j, 9) = IIf(IsNull(rsE.Fields("DeskKelas").Value), "-", rsE.Fields("DeskKelas")) 'KelasBaruM
                            .TextMatrix(j, 28) = 1
                        End If
                    End With
                    rsB.MoveNext
                Next j
            Else
                For j = 1 To rsB.RecordCount
                    With fgData
                        .Rows = rsB.RecordCount + 1
                        .TextMatrix(j, 4) = IIf(IsNull(rsB.Fields(0).Value), "-", rsB.Fields(0))
                        .TextMatrix(j, 5) = IIf(IsNull(rsB.Fields(1).Value), "-", rsB.Fields(1))
                        .TextMatrix(j, 6) = IIf(IsNull(rsB.Fields(2).Value), "-", rsB.Fields(2))
                        .TextMatrix(j, 7) = IIf(IsNull(rsB.Fields(3).Value), "-", rsB.Fields(3))
                        .TextMatrix(j, 8) = IIf(IsNull(rsB.Fields(7).Value), "-", rsB.Fields(7)) 'RuanganTujuanM
                        strSQLy = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsB("NoPendaftaran") & "' and KdRuangan = '" & rsB.Fields("KdRuanganTujuan") & "'"
                        Set rsE = Nothing
                        rsE.Open strSQLy, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsE.EOF = True Then
                            .TextMatrix(j, 9) = "Belum Masuk Ruangan" 'KelasBaruM
                            .TextMatrix(j, 28) = "1"
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsE("KdKelasPel") & "'"
                            Set rsE = Nothing
                            rsE.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(j, 9) = IIf(IsNull(rsE.Fields("DeskKelas").Value), "-", rsE.Fields("DeskKelas")) 'KelasBaruM
                            .TextMatrix(j, 28) = 1
                        End If
                    End With
                    rsB.MoveNext
                Next j
            End If
        End If
        'pasienDipindahkan

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas,NoPendaftaran,NoPakai,KdRuanganTujuan,RuanganTujuan From LaporanSensusHarianRIPasienDiPindahkan_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & dcRuangan.BoundText & "%' ORDER BY RuanganPelayanan"
        Set rsC = Nothing
        rsC.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        If rsC.EOF = False Then
            ipd = rsC.RecordCount

            If ipm >= ipp And ipm >= ipd Then
                For k = 1 To rsC.RecordCount
                    With fgData
                        .Rows = rs.RecordCount + 1
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = "1"
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = 1
                        End If
                    End With
                    rsC.MoveNext
                Next k
            ElseIf ipp >= ipm And ipp >= ipd Then
                For k = 1 To rsC.RecordCount
                    With fgData
                        .Rows = rsB.RecordCount + 1
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = "1"
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = 1
                        End If
                    End With
                    rsC.MoveNext
                Next k
            ElseIf ipd >= ipm And ipd >= ipp Then
                For k = 1 To rsC.RecordCount
                    With fgData
                        .Rows = rsC.RecordCount + 1
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = "1"
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = 1
                        End If
                    End With
                    rsC.MoveNext
                Next k
            End If
        End If

        strSQL = "Select Record,NamaPasien,Kelas,TglSensus,KdKondisiPulang,KdStatusPulang,StatusPulang,Jmlpasien,Tglmasuk,LamaDirawat,RuanganPelayanan From LaporanSensusHarianRIPasienPulang_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & dcRuangan.BoundText & "%' Order By RuanganPelayanan"

        Set rsD = Nothing
        rsD.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        If rsD.EOF = False Then

            ipk = rsD.RecordCount

            If ipm >= ipp And ipm >= ipd And ipm >= ipk Then
                For l = 1 To rsD.RecordCount
                    With fgData
                        .Rows = rs.RecordCount + 1

                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = 1

                    End With
                    rsD.MoveNext
                Next l
            ElseIf ipp >= ipm And ipp >= ipd And ipp >= ipk Then
                For l = 1 To rsC.RecordCount
                    With fgData
                        .Rows = rsB.RecordCount + 1
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = 1
                    End With
                    rsC.MoveNext
                Next l
            ElseIf ipd >= ipm And ipd >= ipp And ipd >= ipk Then
                For l = 1 To rsC.RecordCount
                    With fgData
                        .Rows = rsC.RecordCount + 1
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = 1
                    End With
                    rsC.MoveNext
                Next l

            ElseIf ipk >= ipm And ipk >= ipp And ipk >= ipd Then
                For l = 1 To rsD.RecordCount
                    With fgData
                        .Rows = rsD.RecordCount + 1
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = 1
                    End With
                    rsD.MoveNext
                Next l
            End If
        End If
    Else
        Call subtidakDipilih
    End If
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    'On Error Resume Next
    Dim m As Integer
    Dim pesan As VbMsgBoxResult
    strSQL = "Delete From LaporanSensusHarian_V"
    Call msubRecFO(rs, strSQL)

    strSQLX = ""
    strSQL = ""
    strSQLz = ""
    strSQLa = ""
    strSQLb = ""
    strSQLc = ""
    strSQLd = ""
    strSQLe = ""
    strSQLe = ""
    strSQLf = ""
    strSQLg = ""
    strSQLh = ""
    strSQLj = ""
    strSQLk = ""
    strSQLl = ""

    If optResume.Value = False Then
        If dcRuangan.Text = "" Then
            strCetak2 = "Global"
        End If
        If fgData.Rows - 1 = 0 Then Exit Sub
        With fgData
            For m = 1 To .Rows - 1
                If sp_SimpanLaporanSensus(.TextMatrix(m, 0), IIf(.TextMatrix(m, 1) = "", "-", .TextMatrix(m, 1)), .TextMatrix(m, 2), _
                    .TextMatrix(m, 3), .TextMatrix(m, 4), .TextMatrix(m, 5), _
                    .TextMatrix(m, 6), .TextMatrix(m, 7), .TextMatrix(m, 8), _
                    .TextMatrix(m, 9), .TextMatrix(m, 10), .TextMatrix(m, 11), _
                    .TextMatrix(m, 12), .TextMatrix(m, 13), .TextMatrix(m, 14), _
                    .TextMatrix(m, 15), .TextMatrix(m, 16), .TextMatrix(m, 17), _
                    .TextMatrix(m, 18), .TextMatrix(m, 19), .TextMatrix(m, 20), _
                    .TextMatrix(m, 21), .TextMatrix(m, 22), .TextMatrix(m, 23), _
                    .TextMatrix(m, 24), .TextMatrix(m, 25), .TextMatrix(m, 26), _
                    .TextMatrix(m, 27), .TextMatrix(m, 28)) = False Then Exit Sub

                Next m
            End With
            strSQLz = "Select * from LaporanSensusHarian_V"
            frmCetakSensusHarian.Show
        Else
            strCetak = "resume"

            strSQLX = "union all SELECT Ruangan AS RuanganPelayanan, Kelas, 'b. Penderita Masuk' AS Judul, COUNT(*) AS Jumlah " & _
            "FROM dbo.LaporanSensusHarianRI AS LaporanSensusHarianRIPasienMasuk_V_1 " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglMasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and kdcaramasuk <>'04' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLz = "union all SELECT Ruangan AS RuanganPelayanan, Kelas, 'c. Penderita Pindahan' AS Judul, COUNT(*) AS Jumlah " & _
            "From dbo.LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND Tglmasuk BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and KdCaraMasuk ='04' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLa = "union all SELECT Ruangan AS RuanganPelayanan, Kelas, 'd. Penderita Dirawat'  AS Judul, COUNT(*) AS Jumlah " & _
            "From dbo.LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND Tglmasuk <= '" & Format(dtpAwal, "yyyy/MM/dd 23:59:59") & "' AND (TglKeluar > '" & Format(DateAdd("d", -1, dtpAkhir), "yyyy/MM/dd 23:59:59") & "' or tglkeluar is null) " & _
            "GROUP BY Kelas, Ruangan "

            '  pasien pulang tidak termasuk pasien pindah kamar makanya filter ditambah dengan statuskeluar kamar not in (01)
            strSQLb = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'e. Penderita Pulang' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdStatusKeluar not in ('01') AND KdKondisiPulang NOT IN ('04','05') AND KdStatusPulang IN ('01','10') AND KdRuangan Like '%" & dcRuangan.BoundText & "%' AND Tglpulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLc = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'f. Penderita Referal' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdKondisiPulang IN ('06','02','01','03') AND KdStatusPulang IN ('05','02','03','06','07','04') AND KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLd = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'g. Penderita APS' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdKondisiPulang = '11' AND  KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglKeluar BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLe = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'h. Penderita Lari' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdKondisiPulang <> '11' AND KdStatusPulang = '09' AND  KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLf = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'i. Jml Penderita Keluar hidup' AS Judul, COUNT(*) AS Jumlah From dbo.LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglKeluar between '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and kdkondisipulang not in('04','05','09','10')and KdStatusKeluar <>'01'  " & _
            "GROUP BY Kelas, Ruangan "

            strSQLg = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'j. Penderita Dipindahkan' AS Judul, COUNT(*) AS Jumlah " & _
            "From LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglKeluar BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' and KdStatusKeluar ='01' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLh = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'k. Penderita Mati < 48 Jam' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdKondisiPulang = '04' AND KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLi = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'l. Penderita Mati > 48 Jam' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdKondisiPulang = '05' AND  KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglPulang BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLj = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'm. Jml Penderita Keluar' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND TglKeluar BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
            "GROUP BY Kelas, Ruangan "

            strSQLk = "union all SELECT Ruangan as RuanganPelayanan,Kelas, 'n. Jml Penderita Masih Dirawat' AS Judul, COUNT(*) AS Jumlah " & _
            "FROM LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND tglmasuk <= '" & Format(dtpAwal, "yyyy/MM/dd 23:59:59") & "' AND (TglKeluar > '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' or TglKeluar is null) " & _
            "GROUP BY Kelas, Ruangan "

            strSQLl = "union all SELECT Ruangan AS RuanganPelayanan,Kelas, 'o. Penderita M K pada Hari yang Sama' AS Judul, COUNT(*) AS Jumlah FROM LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND (TglRegistrasi between '" & Format(dtpAwal, "yyyy/MM/dd") & "' and '" & Format(dtpAwal, "yyyy/MM/dd 23:59:59") & "') AND (TglPulang between '" & Format(dtpAkhir, "yyyy/MM/dd") & "' and '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "')  " & _
            "GROUP BY Kelas, Ruangan "

            strSQL = "SELECT Ruangan AS RuanganPelayanan, Kelas, 'a. Penderita Awal' AS Judul, COUNT(*) AS Jumlah " & _
            "From dbo.LaporanSensusHarianRI " & _
            "WHERE KdRuangan Like '%" & dcRuangan.BoundText & "%' AND tglmasuk <= '" & Format(DateAdd("d", -1, dtpAwal), "yyyy/MM/dd 23:59:59") & "' AND (TglKeluar > '" & Format(DateAdd("d", -1, dtpAkhir), "yyyy/MM/dd 23:59:59") & "' or tglkeluar is null) " & _
            "GROUP BY Kelas, Ruangan " & strSQLX & strSQLz & strSQLa & strSQLb & strSQLc & strSQLd & strSQLe & strSQLf & strSQLg & strSQLh & strSQLi & strSQLj & strSQLk & strSQLl

            If fgData.Rows - 1 = 0 Then Exit Sub
            With fgData
                For m = 1 To .Rows - 1
                    If sp_SimpanLaporanSensus(.TextMatrix(m, 0), IIf(.TextMatrix(m, 1) = "", "-", .TextMatrix(m, 1)), .TextMatrix(m, 2), _
                        .TextMatrix(m, 3), .TextMatrix(m, 4), .TextMatrix(m, 5), _
                        .TextMatrix(m, 6), .TextMatrix(m, 7), .TextMatrix(m, 8), _
                        .TextMatrix(m, 9), .TextMatrix(m, 10), .TextMatrix(m, 11), _
                        .TextMatrix(m, 12), .TextMatrix(m, 13), .TextMatrix(m, 14), _
                        .TextMatrix(m, 15), .TextMatrix(m, 16), .TextMatrix(m, 17), _
                        .TextMatrix(m, 18), .TextMatrix(m, 19), .TextMatrix(m, 20), _
                        .TextMatrix(m, 21), .TextMatrix(m, 22), .TextMatrix(m, 23), _
                        .TextMatrix(m, 24), .TextMatrix(m, 25), .TextMatrix(m, 26), _
                        .TextMatrix(m, 27), .TextMatrix(m, 28)) = False Then Exit Sub

                    Next m
                End With
                strSQLz = "Select * from LaporanSensusHarian_V"

                pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
                vLaporan = ""
                If pesan = vbYes Then vLaporan = "Print"

                frmCetakSensusHarian.Show
            End If
            Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    strSQLz = "Delete From LaporanSensusHarian_V"
    Call msubRecFO(dbRst, strSQLz)
    Unload Me
End Sub

Private Sub Form_Load()
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    With Me
        .dtpAwal.Value = Format(Now, "dd MM yyyy 00:00:00")
        .dtpAkhir.Value = Format(Now, "dd MM yyyy 23:59:59")
    End With
    Call subDcSource
    dcRuangan.Text = mstrNamaRuangan
    Call subSetGridClear
End Sub

Private Sub subDcSource()
    strSQL = "select KdRuangan,NamaRuangan From V_RuanganPelayanan where KdInstalasi in ('03', '08', '26') ORDER BY NamaRuangan ASC"
    Call msubDcSource(dcRuangan, rs, strSQL)
'    If rs.EOF = True Then Exit Sub
'    dcRuangan.BoundText = rs(0).Value
End Sub

Private Sub subSetGridClear()
    Dim i As Integer
    With fgData
        .clear
        .Rows = 2
        .Cols = 29

        .ColWidth(0) = 1500 'NoCMPM
        .ColWidth(1) = 2500 'NamaPM
        .ColWidth(2) = 2500 'RuanganPM
        .ColWidth(3) = 1000 'KelasPM

        .ColWidth(4) = 1500 'NoCMPP
        .ColWidth(5) = 3000 'NamaPP
        .ColWidth(6) = 2500 'RuanganAsal
        .ColWidth(7) = 1000 'KelasPP
        .ColWidth(8) = 2500 'RuanganBaru
        .ColWidth(9) = 2000 'KelasBaru

        .ColWidth(10) = 1500 'NoCMPD
        .ColWidth(11) = 3000 'NamaPD
        .ColWidth(12) = 2500 'RuanganAsal
        .ColWidth(13) = 1000 'KelasAsal
        .ColWidth(14) = 2500 'RuanganTujuan
        .ColWidth(15) = 2000 'KelasPD

        .ColWidth(16) = 2000 'TglMasuk
        .ColWidth(17) = 1500 'NoCMPK
        .ColWidth(18) = 3000 'NamaPK
        .ColWidth(19) = 1000 'KelasPK

        .ColWidth(20) = 1000 'Pulang
        .ColWidth(21) = 1000 'Referal
        .ColWidth(22) = 1000 'APS
        .ColWidth(23) = 1000 'Lari
        .ColWidth(24) = 1500 'Mati < 48 Jam
        .ColWidth(25) = 1500 'Mati >=48 Jam
        .ColWidth(26) = 1500 'LamaRawat
        .ColWidth(27) = 1500 'RuanganAkhir
        .ColWidth(28) = 0 'Status

        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter

        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignLeftCenter

        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .ColAlignment(12) = flexAlignLeftCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignLeftCenter
        .ColAlignment(15) = flexAlignLeftCenter

        .ColAlignment(16) = flexAlignCenterCenter
        .ColAlignment(17) = flexAlignCenterCenter
        .ColAlignment(18) = flexAlignLeftCenter
        .ColAlignment(19) = flexAlignCenterCenter

        .ColAlignment(20) = flexAlignCenterCenter
        .ColAlignment(21) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(24) = flexAlignCenterCenter
        .ColAlignment(25) = flexAlignCenterCenter
        .ColAlignment(26) = flexAlignCenterCenter
        .ColAlignment(27) = flexAlignCenterCenter

    End With
    Call subSetGrid
End Sub

Private Sub subSetGrid()
    With fgData
        .TextMatrix(0, 0) = "NoCMPM"
        .TextMatrix(0, 1) = "NamaPM"
        .TextMatrix(0, 2) = "Ruangan"
        .TextMatrix(0, 3) = "KelasPM"

        .TextMatrix(0, 4) = "NoCMPP"
        .TextMatrix(0, 5) = "NamaPP"
        .TextMatrix(0, 6) = "RuanganAsal"
        .TextMatrix(0, 7) = "KelasPP"
        .TextMatrix(0, 8) = "RuanganBaruM"
        .TextMatrix(0, 9) = "KelasBaruM"

        .TextMatrix(0, 10) = "NoCMPD"
        .TextMatrix(0, 11) = "NamaPD"
        .TextMatrix(0, 12) = "RuanganAsalPD"
        .TextMatrix(0, 13) = "KelasAsalPD"
        .TextMatrix(0, 14) = "RuanganTujuanPD"
        .TextMatrix(0, 15) = "KelasPD"

        .TextMatrix(0, 16) = "TglMasuk"
        .TextMatrix(0, 17) = "NoCMPK"
        .TextMatrix(0, 18) = "NamaPK"
        .TextMatrix(0, 19) = "KelasPK"

        .TextMatrix(0, 20) = "Pulang"
        .TextMatrix(0, 21) = "Referal"
        .TextMatrix(0, 22) = "APS"
        .TextMatrix(0, 23) = "Lari"
        .TextMatrix(0, 24) = "Mati < 48 Jam"
        .TextMatrix(0, 25) = "Mati >=48 Jam"
        .TextMatrix(0, 26) = "Lama Rawat"
        .TextMatrix(0, 27) = "RuanganAkhir"
        .TextMatrix(0, 28) = "Status"

    End With
End Sub

Private Function sp_SimpanLaporanSensus(f_NoCMPM As String, f_NamaPM As String, f_Ruangan As String, f_KelasPM As String, _
    f_NoCMPP As String, f_NamaPP As String, f_RuanganAsalPP As String, f_KelasPP As String, f_RuanganBaruPP As String, f_KelasBaruPP As String, _
    f_NoCMPD As String, f_NamaPD As String, f_RuanganAsalPD As String, f_KelasAsalPD As String, f_RuanganTujuanPD As String, f_KelasPD As String, _
    f_TglMasuk As String, f_NoCMPK As String, f_NamaPK As String, f_KelasPK As String, _
    f_Pulang As String, f_Referal As String, f_APS As String, f_Lari As String, f_MatiDown As String, f_MatiUp As String, f_lamarawat As String, _
    f_RuanganAkhir As String, f_status As String) As Boolean
    sp_SimpanLaporanSensus = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_Value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoCMPM", adVarChar, adParamInput, 12, f_NoCMPM)
        .Parameters.Append .CreateParameter("NamaPM", adChar, adParamInput, 50, f_NamaPM)
        .Parameters.Append .CreateParameter("RuanganPelayanan", adChar, adParamInput, 30, f_Ruangan)
        .Parameters.Append .CreateParameter("KelasPM", adChar, adParamInput, 20, f_KelasPM)

        .Parameters.Append .CreateParameter("NoCMPP", adVarChar, adParamInput, 12, f_NoCMPP)
        .Parameters.Append .CreateParameter("NamaPP", adChar, adParamInput, 50, f_NamaPP)
        .Parameters.Append .CreateParameter("RuanganAsalPP", adChar, adParamInput, 30, f_RuanganAsalPP)
        .Parameters.Append .CreateParameter("KelasAsalPP", adChar, adParamInput, 20, f_KelasPP)
        .Parameters.Append .CreateParameter("RuanganBaruPP", adChar, adParamInput, 30, f_RuanganBaruPP)
        .Parameters.Append .CreateParameter("KelasBaruPP", adChar, adParamInput, 20, f_KelasBaruPP)

        .Parameters.Append .CreateParameter("NoCMPD", adVarChar, adParamInput, 12, f_NoCMPD)
        .Parameters.Append .CreateParameter("NamaPD", adChar, adParamInput, 50, f_NamaPD)
        .Parameters.Append .CreateParameter("RuanganAsalPD", adChar, adParamInput, 30, f_RuanganAsalPD)
        .Parameters.Append .CreateParameter("KelasAsalPD", adChar, adParamInput, 20, f_KelasAsalPD)
        .Parameters.Append .CreateParameter("RuanganBaruPD", adChar, adParamInput, 30, f_RuanganTujuanPD)
        .Parameters.Append .CreateParameter("KelasBaruPD", adChar, adParamInput, 20, f_KelasPD)

        .Parameters.Append .CreateParameter("TglMasuk", adChar, adParamInput, 20, Format(f_TglMasuk, "yyyy/mm/dd"))
        .Parameters.Append .CreateParameter("NoCMPK", adVarChar, adParamInput, 12, f_NoCMPK)
        .Parameters.Append .CreateParameter("NamaPK", adChar, adParamInput, 50, f_NamaPK)
        .Parameters.Append .CreateParameter("KelasPK", adChar, adParamInput, 20, f_KelasPK)

        .Parameters.Append .CreateParameter("Pulang", adChar, adParamInput, 1, f_Pulang)
        .Parameters.Append .CreateParameter("Referal", adChar, adParamInput, 1, f_Referal)
        .Parameters.Append .CreateParameter("Aps", adChar, adParamInput, 1, f_APS)
        .Parameters.Append .CreateParameter("Lari", adChar, adParamInput, 1, f_Lari)
        .Parameters.Append .CreateParameter("MatiDown", adChar, adParamInput, 1, f_MatiDown)
        .Parameters.Append .CreateParameter("MatiUp", adChar, adParamInput, 1, f_MatiUp)
        .Parameters.Append .CreateParameter("LamaDirawat", adChar, adParamInput, 3, f_lamarawat)
        .Parameters.Append .CreateParameter("RuanganAkhir", adChar, adParamInput, 30, f_RuanganAkhir)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 3, (f_status))

        .ActiveConnection = dbConn
        .CommandText = "dbo.LaporanSensusHarian_VAdd"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("return_value").Value = 0) Then
            sp_SimpanLaporanSensus = False
            MsgBox "Ada kesalahan dalam Penyimpanan laporan Sensus", vbExclamation, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
    End With
End Function

Private Sub subtidakDipilih()
    Call subSetGridClear
    strSQLm = "select KdRuangan,NamaRuangan From V_RuanganPelayanan where KdInstalasi='03' ORDER BY NamaRuangan ASC"
    Call msubRecFO(rsa, strSQLm)

    p = 1
    q = 1

    For o = 1 To rsa.RecordCount

        'pasien masuk

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas From LaporanSensusHarianRIPasienMasuk_V " & _
        "where  TglSensus BETWEEN '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & rsa(0).Value & "%' ORDER BY RuanganPelayanan"

        Set rs = Nothing
        rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        ipm = 0

        If rs.EOF = False Then

            ipm = rs.RecordCount
            For i = p To p + rs.RecordCount - 1
                With fgData
                    .Rows = p + rs.RecordCount
                    .TextMatrix(i, 0) = IIf(IsNull(rs.Fields(0).Value), "-", rs.Fields(0))
                    .TextMatrix(i, 1) = IIf(IsNull(rs.Fields(1).Value), "-", rs.Fields(1))
                    .TextMatrix(i, 2) = IIf(IsNull(rs.Fields(2).Value), "-", rs.Fields(2))
                    .TextMatrix(i, 3) = IIf(IsNull(rs.Fields(3).Value), "-", rs.Fields(3))
                    .TextMatrix(i, 28) = q
                End With
                rs.MoveNext
            Next i
        End If
        'Pasien Pindahan

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas,NoPendaftaran,NoPakai,KdRuanganTujuan,RuanganTujuan From LaporanSensusHarianRIPasienDiPindahkan_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganTujuan Like '%" & rsa(0).Value & "%' ORDER BY RuanganTujuan"
        Set rsB = Nothing
        rsB.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        ipp = 0
        If rsB.EOF = False Then
            ipp = rsB.RecordCount

            If ipm >= ipp Then
                For j = p To p + rsB.RecordCount - 1
                    With fgData
                        .Rows = rs.RecordCount + p
                        .TextMatrix(j, 4) = IIf(IsNull(rsB.Fields(0).Value), "-", rsB.Fields(0))
                        .TextMatrix(j, 5) = IIf(IsNull(rsB.Fields(1).Value), "-", rsB.Fields(1))
                        .TextMatrix(j, 6) = IIf(IsNull(rsB.Fields(2).Value), "-", rsB.Fields(2))
                        .TextMatrix(j, 7) = IIf(IsNull(rsB.Fields(3).Value), "-", rsB.Fields(3))
                        .TextMatrix(j, 8) = IIf(IsNull(rsB.Fields(7).Value), "-", rsB.Fields(7))  'RuanganTujuanM
                        strSQLy = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsB("NoPendaftaran") & "' and KdRuangan = '" & rsB.Fields("KdRuanganTujuan") & "'"
                        Set rsE = Nothing
                        rsE.Open strSQLy, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsE.EOF = True Then
                            .TextMatrix(j, 9) = "Belum Masuk Ruangan"
                            .TextMatrix(j, 28) = q
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsE("KdKelasPel") & "'"
                            Set rsE = Nothing
                            rsE.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(j, 9) = IIf(IsNull(rsE.Fields("DeskKelas").Value), "-", rsE.Fields("DeskKelas")) 'KelasBaruM
                            .TextMatrix(j, 28) = q
                        End If
                    End With
                    rsB.MoveNext
                Next j
            Else
                For j = p To p + rsB.RecordCount - 1
                    With fgData
                        .Rows = rsB.RecordCount + p
                        .TextMatrix(j, 4) = IIf(IsNull(rsB.Fields(0).Value), "-", rsB.Fields(0))
                        .TextMatrix(j, 5) = IIf(IsNull(rsB.Fields(1).Value), "-", rsB.Fields(1))
                        .TextMatrix(j, 6) = IIf(IsNull(rsB.Fields(2).Value), "-", rsB.Fields(2))
                        .TextMatrix(j, 7) = IIf(IsNull(rsB.Fields(3).Value), "-", rsB.Fields(3))
                        .TextMatrix(j, 8) = IIf(IsNull(rsB.Fields(7).Value), "-", rsB.Fields(7)) 'RuanganTujuanM
                        strSQLy = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsB("NoPendaftaran") & "' and KdRuangan = '" & rsB.Fields("KdRuanganTujuan") & "'"
                        Set rsE = Nothing
                        rsE.Open strSQLy, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsE.EOF = True Then
                            .TextMatrix(j, 9) = "Belum Masuk Ruangan" 'KelasBaruM
                            .TextMatrix(j, 28) = q
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsE("KdKelasPel") & "'"
                            Set rsE = Nothing
                            rsE.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(j, 9) = IIf(IsNull(rsE.Fields("DeskKelas").Value), "-", rsE.Fields("DeskKelas")) 'KelasBaruM
                            .TextMatrix(j, 28) = q
                        End If
                    End With
                    rsB.MoveNext
                Next j
            End If
        End If
        'pasienDipindahkan

        strSQL = "Select Record,NamaPasien,RuanganPelayanan,Kelas,NoPendaftaran,NoPakai,KdRuanganTujuan,RuanganTujuan From LaporanSensusHarianRIPasienDiPindahkan_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & rsa(0).Value & "%' ORDER BY RuanganPelayanan"
        Set rsC = Nothing
        rsC.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        ipd = 0
        If rsC.EOF = False Then
            ipd = rsC.RecordCount

            If ipm >= ipp And ipm >= ipd Then
                For k = p To p + rsC.RecordCount - 1
                    With fgData
                        .Rows = rs.RecordCount + p
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = q
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = q
                        End If
                    End With
                    rsC.MoveNext
                Next k
            ElseIf ipp >= ipm And ipp >= ipd Then
                For k = p To p + rsC.RecordCount - 1
                    With fgData
                        .Rows = rsB.RecordCount + p
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = q
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = q
                        End If
                    End With
                    rsC.MoveNext
                Next k
            ElseIf ipd >= ipm And ipd >= ipp Then
                For k = p To p + rsC.RecordCount - 1
                    With fgData
                        .Rows = rsC.RecordCount + p
                        .TextMatrix(k, 10) = IIf(IsNull(rsC.Fields(0).Value), "-", rsC.Fields(0))
                        .TextMatrix(k, 11) = IIf(IsNull(rsC.Fields(1).Value), "-", rsC.Fields(1))
                        .TextMatrix(k, 12) = IIf(IsNull(rsC.Fields(2).Value), "-", rsC.Fields(2))
                        .TextMatrix(k, 13) = IIf(IsNull(rsC.Fields(3).Value), "-", rsC.Fields(3))
                        .TextMatrix(k, 14) = IIf(IsNull(rsC.Fields(7).Value), "-", rsC.Fields(7)) 'Ruangan Tujuan
                        strSQLX = "Select KdKelasPel from PemakaianKamar where NoPendaftaran= '" & rsC("NoPendaftaran") & "' and KdRuangan = '" & rsC.Fields("KdRuanganTujuan") & "'"
                        Set rsB = Nothing
                        rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                        If rsB.EOF = True Then
                            .TextMatrix(k, 15) = "Belum Masuk Ruangan"
                            .TextMatrix(k, 28) = q
                        Else
                            strSQLX = "Select DeskKelas from KelasPelayanan where KdKelas='" & rsB("KdKelasPel") & "'"
                            Set rsB = Nothing
                            rsB.Open strSQLX, dbConn, adOpenForwardOnly, adLockReadOnly
                            .TextMatrix(k, 15) = IIf(IsNull(rsB.Fields("DeskKelas").Value), "-", rsB.Fields("DeskKelas"))
                            .TextMatrix(k, 28) = q
                        End If
                    End With
                    rsC.MoveNext
                Next k
            End If
        End If
        'pasien Keluar
        strSQL = "Select Record,NamaPasien,Kelas,TglSensus,KdKondisiPulang,KdStatusPulang,StatusPulang,Jmlpasien,Tglmasuk,LamaDirawat,RuanganPelayanan From LaporanSensusHarianRIPasienKeluar_V " & _
        "where  TglSensus BETWEEN " & _
        " '" & Format(dtpAwal, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(dtpAkhir, "yyyy/MM/dd 23:59:59") & "' " & _
        " AND KdRuanganPelayanan Like '%" & rsa(0).Value & "%' Order By RuanganPelayanan"
        Set rsD = Nothing
        rsD.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly

        ipk = 0
        If rsD.EOF = False Then

            ipk = rsD.RecordCount

            If ipm >= ipp And ipm >= ipd And ipm >= ipk Then
                For l = p To p + rsD.RecordCount - 1
                    With fgData
                        .Rows = rs.RecordCount + p

                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = q

                    End With
                    rsD.MoveNext
                Next l
            ElseIf ipp >= ipm And ipp >= ipd And ipp >= ipk Then
                For l = p To p + rsC.RecordCount - 1
                    With fgData
                        .Rows = rsB.RecordCount + p
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = q
                    End With
                    rsC.MoveNext
                Next l
            ElseIf ipd >= ipm And ipd >= ipp And ipd >= ipk Then
                For l = p To p + rsC.RecordCount - 1
                    With fgData
                        .Rows = rsC.RecordCount + p
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8)) 'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = q
                    End With
                    rsC.MoveNext
                Next l

            ElseIf ipk >= ipm And ipk >= ipp And ipk >= ipd Then

                For l = p To p + rsD.RecordCount - 1
                    With fgData

                        .Rows = p + rsD.RecordCount
                        .TextMatrix(l, 16) = IIf(IsNull(rsD.Fields(3).Value), "-", rsD.Fields(8))  'Tglmasuk
                        .TextMatrix(l, 17) = IIf(IsNull(rsD.Fields(0).Value), "-", rsD.Fields(0))
                        .TextMatrix(l, 18) = IIf(IsNull(rsD.Fields(1).Value), "-", rsD.Fields(1))
                        .TextMatrix(l, 19) = IIf(IsNull(rsD.Fields(2).Value), "-", rsD.Fields(2))
                        If rsD(4).Value = "01" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "02" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "03" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "06" Then
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))  'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "08" Then 'Referal
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "11" Then 'APS
                            If rsD(5).Value = "01" Or rsD(5).Value = "10" Then
                                .TextMatrix(l, 20) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Pulang
                            ElseIf rsD(5).Value = "02" Or rsD(5).Value = "03" Or rsD(5).Value = "04" Or rsD(5).Value = "05" Or rsD(5).Value = "06" Or rsD(5).Value = "07" Then
                                .TextMatrix(l, 21) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Referal
                            ElseIf rsD(5).Value = "08" Then
                                .TextMatrix(l, 22) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'APS
                            ElseIf rsD(5).Value = "09" Then
                                .TextMatrix(l, 23) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7)) 'Lari
                            End If
                        ElseIf rsD(4).Value = "04" Then 'Mati < 48 Jam
                            .TextMatrix(l, 24) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        ElseIf rsD(4).Value = "05" Then 'Mati >=48 Jam
                            .TextMatrix(l, 25) = IIf(IsNull(rsD.Fields(7).Value), "-", rsD.Fields(7))
                        End If
                        .TextMatrix(l, 26) = IIf(IsNull(rsD.Fields(9).Value), "-", rsD.Fields(9))
                        .TextMatrix(l, 27) = IIf(IsNull(rsD.Fields(10).Value), "-", rsD.Fields(10))
                        .TextMatrix(l, 28) = q
                    End With
                    rsD.MoveNext
                Next l
            End If
        End If
        rsa.MoveNext
        p = fgData.Rows
        q = q + 1
    Next o
End Sub

