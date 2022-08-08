VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPeriodeIndikatorRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "FrmPeriodeIndikatorRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9975
   Begin VB.Frame Frame4 
      Caption         =   "Jumlah    Pria                   Wanita            Total"
      Height          =   735
      Left            =   5820
      TabIndex        =   13
      Top             =   6510
      Visible         =   0   'False
      Width           =   4155
      Begin VB.TextBox txtJmlTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2970
         TabIndex        =   16
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtJmlWanita 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1890
         TabIndex        =   15
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtJmlPria 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   810
         TabIndex        =   14
         Top             =   240
         Width           =   1000
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgdata 
      Height          =   4425
      Left            =   0
      TabIndex        =   11
      Top             =   2100
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7805
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Instalasi Pelayanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5010
      TabIndex        =   8
      Top             =   930
      Width           =   4935
      Begin VB.ComboBox cboKriteria 
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
         ItemData        =   "FrmPeriodeIndikatorRS.frx":08CA
         Left            =   150
         List            =   "FrmPeriodeIndikatorRS.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   3645
      End
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari"
         Height          =   555
         Left            =   3990
         TabIndex        =   12
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kriteria"
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
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   300
         Width           =   555
      End
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
      Height          =   1155
      Left            =   0
      TabIndex        =   3
      Top             =   930
      Width           =   4995
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
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
         Format          =   51052547
         CurrentDate     =   37956
      End
      Begin MSComCtl2.DTPicker DTPickerAkhir 
         Height          =   330
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
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
         Format          =   51052547
         CurrentDate     =   37956
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Akhir"
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
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame2 
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
      TabIndex        =   0
      Top             =   6510
      Width           =   5805
      Begin VB.CommandButton cmdgrafik 
         Caption         =   "Grafik"
         Height          =   375
         Left            =   2070
         TabIndex        =   9
         Top             =   240
         Width           =   1665
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
         Height          =   375
         Left            =   210
         TabIndex        =   2
         Top             =   240
         Width           =   1665
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutu&p"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3900
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   -210
      Picture         =   "FrmPeriodeIndikatorRS.frx":08FA
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "FrmPeriodeIndikatorRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iRowNow As Integer
Dim rsstatusPasien As ADODB.recordset
Dim rsstatusPasien1 As ADODB.recordset
Dim iRowNow2 As Integer

Private Sub cmdCari_Click()
Dim intJmlRow As Integer
Dim intJmlPria As Integer
Dim intJmlWanita As Integer
Dim intJmlTotal As Integer

    Call subSetGrid
    'u/ mempercepat
    fgdata.Visible = False: MousePointer = vbHourglass
    
    If cboKriteria.Text = "Per Ruangan" Then
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM v_S_RekapIndikatorPlyn " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).Value = 0 Then
        fgdata.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada Data"
'        MsgBox "Tidak ada data antara tanggal  '" & Format(DTPickerAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbInformation, "Validasi"
        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        Exit Sub
    End If
    
    intJmlRow = rs("JmlRow").Value
    
    strSQL = "SELECT namaruangan, " & _
        " round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR From v_S_RekapIndikatorPlyn " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " GROUP BY namaruangan"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    'Tambahkan jumlah row dengan jumlah subtotal
    'intJmlRow = rs.RecordCount 'intJmlRow + rs.RecordCount
'    iRowNow = 0
    'u/ menampilkan yang di group by
    
    With fgdata
        'jml baris akhir
        .Rows = intJmlRow '+ 1
        While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("namaruangan").Value
                .TextMatrix(iRowNow, 2) = rs("JmlTOI").Value & " " & "hari"
                .TextMatrix(iRowNow, 3) = rs("JmlBOR").Value & " " & "%"
                .TextMatrix(iRowNow, 4) = rs("JmlBTO").Value & " " & "kali"
                .TextMatrix(iRowNow, 5) = rs("JmlLOS").Value & " " & "hari"
                .TextMatrix(iRowNow, 6) = rs("JmlGDR").Value & " " & "‰ "
                .TextMatrix(iRowNow, 7) = rs("JmlNDR").Value & " " & "‰ "
            
            'tampilan Black & White
            For i = 1 To .Cols - 1
                .Col = i
                .Row = iRowNow
                .CellBackColor = vbBlackness
                .CellForeColor = vbWhite
'                 If .Col = 1 Then .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1): .CellBackColor = vbWhite: .CellForeColor = vbBlack
'                .RowHeight(.Row) = 300
                .CellFontBold = True
            Next
                      
            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2
          
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next
        
        .Visible = True: MousePointer = vbNormal
    End With


    ElseIf cboKriteria.Text = "Per Kelas" Then
    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM v_S_RekapIndikatorPlyn " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).Value = 0 Then
        fgdata.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada Data"
'        MsgBox "Tidak ada data antara tanggal  '" & Format(DTPickerAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbInformation, "Validasi"
        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        Exit Sub
    End If
    
    intJmlRow = rs("JmlRow").Value
    
    strSQL = "SELECT DeskKelas, " & _
        " round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR From v_S_RekapIndikatorPlyn " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " & _
        " GROUP BY DeskKelas"
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    'Tambahkan jumlah row dengan jumlah subtotal
    'intJmlRow = rs.RecordCount 'intJmlRow + rs.RecordCount
'    iRowNow = 0
    'u/ menampilkan yang di group by
    
    With fgdata
        'jml baris akhir
        .Rows = intJmlRow '+ 1
        While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("DeskKelas").Value
                .TextMatrix(iRowNow, 2) = rs("JmlTOI").Value & " " & "hari"
                .TextMatrix(iRowNow, 3) = rs("JmlBOR").Value & " " & "%"
                .TextMatrix(iRowNow, 4) = rs("JmlBTO").Value & " " & "kali"
                .TextMatrix(iRowNow, 5) = rs("JmlLOS").Value & " " & "hari"
                .TextMatrix(iRowNow, 6) = rs("JmlGDR").Value & " " & "‰"
                .TextMatrix(iRowNow, 7) = rs("JmlNDR").Value & " " & "‰"
            
            'tampilan Black & White
            For i = 1 To .Cols - 1
                .Col = i
                .Row = iRowNow
                .CellBackColor = vbBlackness
                .CellForeColor = vbWhite
'                 If .Col = 1 Then .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1): .CellBackColor = vbWhite: .CellForeColor = vbBlack
'                .RowHeight(.Row) = 300
                .CellFontBold = True
            Next
                      
            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2
          
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next
        
        .Visible = True: MousePointer = vbNormal
    End With

    ElseIf cboKriteria.Text = "Semua" Then
'    'Hitung jumlah row dari data yang hendak ditampilkan
    strSQL = "SELECT COUNT(tglhitung) AS JmlRow " & _
        " FROM RekapitulasiIndikatorPelayananRS " & _
        " WHERE tglhitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    'jika tidak ada data
    If rs(0).Value = 0 Then
        fgdata.Visible = True: MousePointer = vbNormal
        MsgBox "Tidak ada Data"
'        MsgBox "Tidak ada data antara tanggal  '" & Format(DTPickerAwal.Value, "dd - MMMM - yyyy") & "' dan '" & Format(dtpTglAkhir.Value, "dd - MMMM - yyyy") & "' ", vbInformation, "Validasi"
'        txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
        Exit Sub
    End If
    
    intJmlRow = rs("JmlRow").Value
    
    
    strSQL = "SELECT  round(AVG(JmlTOI),2) AS JmlTOI,round(AVG(JmlBOR),2) AS JmlBOR," & _
        " round(AVG(JmlBTO),2) AS JmlBTO, round(AVG(JmlLOS),2)AS JmlLOS," & _
        " round(AVG(JmlGDR),2) AS JmlGDR, round(AVG(JmlNDR),2) AS JmlNDR From RekapitulasiIndikatorPelayananRS " & _
        " WHERE TglHitung BETWEEN " & _
        " '" & Format(DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND " & _
        " '" & Format(DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' "
        
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    
    'Tambahkan jumlah row dengan jumlah subtotal
    'intJmlRow = rs.RecordCount 'intJmlRow + rs.RecordCount
'    iRowNow = 0
    'u/ menampilkan yang di group by
    
    With fgdata
        'jml baris akhir
        .Rows = intJmlRow '+ 1
        While rs.EOF = False
                'baris u/ sub total
                iRowNow = iRowNow + 1
                .TextMatrix(iRowNow, 1) = rs("JmlTOI").Value & " " & "hari"
                .TextMatrix(iRowNow, 2) = rs("JmlBOR").Value & " " & "%"
                .TextMatrix(iRowNow, 3) = rs("JmlBTO").Value & " " & "kali"
                .TextMatrix(iRowNow, 4) = rs("JmlLOS").Value & " " & "hari"
                .TextMatrix(iRowNow, 5) = rs("JmlGDR").Value & " " & "‰"
                .TextMatrix(iRowNow, 6) = rs("JmlNDR").Value & " " & "‰"
            
'            tampilan Black & White
            For i = 1 To .Cols - 1
                .Col = i
                .Row = iRowNow
                .CellBackColor = vbBlackness
                .CellForeColor = vbWhite
'                 If .Col = 1 Then .TextMatrix(.Row, 1) = .TextMatrix(.Row, 1): .CellBackColor = vbWhite: .CellForeColor = vbBlack
'                .RowHeight(.Row) = 300
                .CellFontBold = True
            Next
                      
            rs.MoveNext
        Wend
        'banyak baris berdasarkan irownow
        .Rows = iRowNow + 2
          
        .Col = 1
        For i = 1 To .Rows - 1
            .Row = i
            .CellFontBold = True
        Next
        
        .Visible = True: MousePointer = vbNormal
    End With

End If
'End If
End Sub

Private Sub cmdCetak_Click()
If cboKriteria.Text = "Semua" Then
    strSQL = "SELECT AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM RekapitulasiIndikatorPelayananRS " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "Semua"
    frmUtilitasRS2.Show
'    frmUtilitasRS2.Caption = "Indikator Pelayanan RS"
'    Unload Me
ElseIf (cboKriteria.Text = "Per Ruangan") Then
    strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" _
            & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY NamaRuangan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "PerRuangan"
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
'    Unload Me
ElseIf (cboKriteria.Text = "Per Kelas") Then
    strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" _
            & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
            & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY DeskKelas"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "PerKelas"
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
'    Unload Me
End If
End Sub

Private Sub cmdgrafik_Click()
If cboKriteria.Text = "Semua" Then
    strSQL = "SELECT AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
        & "FROM RekapitulasiIndikatorPelayananRS " _
        & "WHERE TglHitung BETWEEN '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAwal, "yyyy/MM/dd 00:00:00") & "' AND '" _
        & Format(FrmPeriodeIndikatorRS.DTPickerAkhir, "yyyy/MM/dd 23:59:59") & "' "
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "GrafikSemua"
    frmUtilitasRS2.Show
'    frmUtilitasRS2.Caption = "Indikator Pelayanan RS"
'    Unload Me
ElseIf (cboKriteria.Text = "Per Ruangan") Then
    strSQL = "SELECT NamaRuangan AS Ruangan,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY NamaRuangan"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "GrafikPerRuangan"
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
'    Unload Me
ElseIf (cboKriteria.Text = "Per Kelas") Then
    strSQL = "SELECT DeskKelas AS Kelas,AVG(JmlTOI) AS TOI,AVG(JmlBOR) AS BOR,AVG(JmlBTO) AS BTO,AVG(JmlLOS) AS LOS,AVG(JmlGDR) AS GDR,AVG(JmlNDR) AS NDR " _
            & "FROM dbo.v_S_RekapIndikatorPlyn " _
            & "WHERE TglHitung BETWEEN '" & Format(FrmPeriodeIndikatorRS.DTPickerAwal.Value, "yyyy/MM/dd 00:00:00") & "' AND '" & Format(FrmPeriodeIndikatorRS.DTPickerAkhir.Value, "yyyy/MM/dd 23:59:59") & "' " _
            & "GROUP BY DeskKelas"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada data", vbCritical, "Validasi"
        cmdCetak.Enabled = True
        Exit Sub
    End If
    cetak = "GrafikPerkelas"
    frmUtilitasRS.Show
'    frmUtilitasRS.Caption = "Indikator Pelayanan RS"
'    Unload Me
End If
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    txtJmlPria = "0": txtJmlTotal = "0": txtJmlWanita = "0"
    
    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
    
'    Set dbRec = New ADODB.Recordset
'    dbRec.Open " SELECT     KdInstalasi, NamaInstalasi " _
'             & " FROM         Instalasi where kdinstalasi <> '06'", dbConn, adOpenDynamic, adLockOptimistic
'
'    While dbRec.EOF = False
'        CboInstalasi.AddItem dbRec.Fields(0).Value & " - " & dbRec.Fields(1).Value
'        dbRec.MoveNext
'    Wend

    Call subSetGrid
'    Call SetText
End Sub

Private Sub subSetGrid()
If cboKriteria.Text = "Per Ruangan" Then
    With fgdata
        .Visible = False
        .Clear
        .Cols = 8
        .Rows = 2
        .Row = 0
        
        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next
        
        .MergeCells = 1
        .MergeCol(1) = True
        
        .TextMatrix(0, 1) = "Ruang Pelayanan"
        .TextMatrix(0, 2) = "TOI"
        .TextMatrix(0, 3) = "BOR"
        .TextMatrix(0, 4) = "BTO"
        .TextMatrix(0, 5) = "LOS"
        .TextMatrix(0, 6) = "GDR"
        .TextMatrix(0, 7) = "NDR"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        .ColWidth(7) = 1100
        
        .Visible = True
        iRowNow = 0
    End With
ElseIf cboKriteria.Text = "Per Kelas" Then
    With fgdata
        .Visible = False
        .Clear
        .Cols = 8
        .Rows = 2
        .Row = 0
        
        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next
        
        .MergeCells = 1
        .MergeCol(1) = True
        
        .TextMatrix(0, 1) = "Kelas Pelayanan"
        .TextMatrix(0, 2) = "TOI"
        .TextMatrix(0, 3) = "BOR"
        .TextMatrix(0, 4) = "BTO"
        .TextMatrix(0, 5) = "LOS"
        .TextMatrix(0, 6) = "GDR"
        .TextMatrix(0, 7) = "NDR"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2850
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        .ColWidth(7) = 1100
        
        .Visible = True
        iRowNow = 0
    End With

ElseIf cboKriteria.Text = "Semua" Then
    With fgdata
        .Visible = False
        .Clear
        .Cols = 7
        .Rows = 2
        .Row = 0
        
        For i = 1 To .Cols - 1
            .Col = i
            .CellFontBold = True
            .RowHeight(0) = 300
            .CellAlignment = flexAlignCenterCenter
        Next
        
        .MergeCells = 1
        .MergeCol(1) = True
        
        .TextMatrix(0, 1) = "TOI"
        .TextMatrix(0, 2) = "BOR"
        .TextMatrix(0, 3) = "BTO"
        .TextMatrix(0, 4) = "LOS"
        .TextMatrix(0, 5) = "GDR"
        .TextMatrix(0, 6) = "NDR"
        
        .ColWidth(0) = 1100
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        
        
        .Visible = True
        iRowNow = 0
    End With

End If
'End If
End Sub
