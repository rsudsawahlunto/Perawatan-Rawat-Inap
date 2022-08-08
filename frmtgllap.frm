VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtgllap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Buku Register Pasien"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmtgllap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5550
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5535
      Begin VB.CommandButton cmdBatal 
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
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   1575
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
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode Laporan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   5535
      Begin MSComCtl2.DTPicker DTPickerAkhir 
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   59572227
         CurrentDate     =   37972
      End
      Begin MSComCtl2.DTPicker DTPickerAwal 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM, yyyy"
         Format          =   59572227
         CurrentDate     =   37972
      End
      Begin VB.Label Label3 
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
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
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
         Left            =   2640
         TabIndex        =   7
         Top             =   682
         Width           =   255
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
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   0
      Picture         =   "frmtgllap.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmtgllap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
     Unload Me
End Sub

Private Sub cmdCetak_Click()
  
    If rs.State = 1 Then rs.Close
    strSQL = "SELECT * " _
             & "FROM V_RegisterPasienRI where kdruangan='" & strNKdRuangan & "' and TglPendaftaran BETWEEN '" & Format(DTPickerAwal, "yyyy/mm/dd 00:00:00") & "' AND '" & Format(DTPickerAkhir, "yyyy/mm/dd 23:59:59") & "'"
    
    rs.Open strSQL, dbConn, , adLockOptimistic
    If rs.EOF Then
        MsgBox "Data Tidak Ada", vbInformation, "Informasi"
        Exit Sub
    End If
    FrmLapHarianRI.Show
End Sub

Private Sub Form_Load()
     Call centerForm(Me, MDIUtama)

    With Me
        .DTPickerAwal.Value = Now
        .DTPickerAkhir.Value = Now
    End With
End Sub
