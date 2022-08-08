Attribute VB_Name = "modFungsiDB"
Public fRS  As New ADODB.recordset
Public fRS2 As New ADODB.recordset
Public fQuery As String
Public fQuery2 As String
'Konversi dari SP: Add_TempHargaKomponen
Public Function f_AddTempHargaKomponen(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Currency, fJmlPelayanan As Integer, fStatusCito As String, fIdPegawai As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As String
    Dim fIdPegawai3 As String
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdRuanganAsal As String
    Dim fNoLab_Rad As String
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdRuanganAsal = fRS("KdRuanganAsal").Value
    
    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then
        fIdPegawai1 = fRS("IdPegawai").Value
        fIdPegawai2 = IIf(IsNull(fRS("IdPegawai2").Value), "", fRS("IdPegawai2").Value)
        fIdPegawai3 = IIf(IsNull(fRS("IdPegawai3").Value), "", fRS("IdPegawai3").Value)
        fTarifKelasPenjaminDB = fRS("TarifKelasPenjamin").Value
        fJmlHutangPenjaminDB = fRS("JmlHutangPenjamin").Value
        fJmlTanggunganRSDB = fRS("JmlTanggunganRS").Value
        fJmlPembebasanDB = fRS("JmlPembebasan").Value
    Else
        fIdPegawai1 = ""
        fIdPegawai2 = ""
        fIdPegawai3 = ""
        fTarifKelasPenjaminDB = ""
        fJmlHutangPenjaminDB = ""
        fJmlTanggunganRSDB = ""
        fJmlPembebasanDB = ""
    End If
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai1 = fRS("KdJenisPegawai").Value Else fKdJenisPegawai1 = ""
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai2 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai2 = fRS("KdJenisPegawai").Value Else fKdJenisPegawai2 = ""
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai3 & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = False Then fKdJenisPegawai3 = fRS("KdJenisPegawai").Value Else fKdJenisPegawai3 = ""
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    If fKdJenisPegawai1 = "001" Then
        fIdDokter = fIdPegawai
    Else
        fIdDokter = ""
    End If
    Set fRS = Nothing
    fQuery = "select KdPelayananRS from ConvertPelayananToJasaDokter where KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
    Else
        If (fIdDokter = "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14')"
        End If
        If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14')"
        End If
        If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
        End If
        If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdDokter <> "") Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
        End If
    End If
    
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponen = fRS("KdKomponen").Value
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = False Then
            fHarga = fRS2("Harga").Value
        Else
            fHarga = 0
        End If
        
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlTanggunganRSDB
        End If
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null,'" & fIdPegawai1 & "'," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null,'" & fIdPegawai2 & "'," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "'," & fHarga & "," & fJmlPelayanan & ", null,'" & fIdPegawai3 & "'," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
            End If
        Else
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "04" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai2 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
            If fKdKomponen = "14" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai3 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk is null"
            End If
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
        fRS.MoveNext
    Wend
    
    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fKdKomponenTarifTotal = "12"
    Else
        fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").Value
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "','" & fIdPegawai1 & "','" & fIdPegawai2 & "','" & fIdPegawai3 & "', 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    fTarifTotal = fRS("Harga").Value
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "'," & fTarifTotal & "," & fJmlPelayanan & ", null,'" & fIdPegawai1 & "'," & fJmlHutangPenjaminDB & "," & fJmlTanggunganRSDB & "," & fJmlPembebasanDB & ",null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifTotal & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPenjaminDB & ",JmlTanggunganRS=" & fJmlTanggunganRSDB & ",JmlPembebasan=" & fJmlPembebasanDB & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk is null"
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    'end Tarif Total
    
    'begin Tarif Cito
    If fStatusCito = "1" Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fKdKomponenTarifCito = "07"
        Else
            fKdKomponenTarifCito = fRS("KdKomponenTarifCito").Value
        End If
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlTanggunganRSDB
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "'," & fTarifCito & "," & fJmlPelayanan & ", null,'" & fIdPegawai1 & "'," & fJmlHutangPerKomp & "," & fJmlTanggunganPerKomp & "," & fJmlPembebasanPerKomp & ",null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifCito & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin=" & fJmlHutangPerKomp & ",JmlTanggunganRS=" & fJmlTanggunganPerKomp & ",JmlPembebasan=" & fJmlPembebasanPerKomp & " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk is null"
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
    End If
    'end Tarif Cito
End Function
'fungsi ini tidak berlaku untuk RSU Haji
'Konversi dari SP: Add_TempHargaKomponenForIBS
Public Function f_AddTempHargaKomponenForIBS(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fJmlPelayanan As Integer)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As String
    Dim fIdPegawai3 As String
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fTarifCito As Currency
    Dim fKdRuanganAsal As String
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    
    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan,TarifCito from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPegawai1 = fRS("IdPegawai").Value
    fIdPegawai2 = fRS("IdPegawai2").Value
    fIdPegawai3 = fRS("IdPegawai3").Value
    fTarifKelasPenjaminDB = fRS("TarifKelasPenjamin").Value
    fJmlHutangPenjaminDB = fRS("JmlHutangPenjamin").Value
    fJmlTanggunganRSDB = fRS("JmlTanggunganRS").Value
    fJmlPembebasanDB = fRS("JmlPembebasan").Value
    fTarifCito = fRS("TarifCito").Value
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai1 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai2 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai2 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai3 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai3 = fRS("KdJenisPegawai").Value
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    Set fRS = Nothing
    If (fIdPegawai1 = "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14','20')"
    End If
    If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14','20')"
    End If
    If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
    End If
    If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdPegawai1 <> "") Then
        fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
    End If
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
        fKdKomponen = fRS("KdKomponen").Value
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
        Call msubRecFO(fRS2, fQuery2)
        fHarga = fRS2("Harga").Value
        If fHarga = "" Then
            fHarga = 0
        End If
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlTanggunganRSDB
        End If
        If fJmlHutangPerKomp = "" Then fJmlHutangPerKomp = 0
        If fJmlTanggunganPerKomp = "" Then fJmlTanggunganPerKomp = 0
        If fKdKomponen = "04" And fIdPegawai2 = "" Then
            fKdKomponen = "26"
        End If
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
            End If
            If fKdKomponen = "04" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai2 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
            End If
            If fKdKomponen = "14" Then
                fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai3 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
            End If
        Else
            If fKdKomponen <> "04" And fKdKomponen <> "14" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
            End If
            If fKdKomponen = "04" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai2 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
            End If
            If fKdKomponen = "14" Then
               fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai3 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
            End If
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
        fRS.MoveNext
    Wend
    
    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fKdKomponenTarifTotal = "12"
    Else
        fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").Value
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "','" & fIdPegawai1 & "','" & fIdPegawai2 & "','" & fIdPegawai3 & "', 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    fTarifTotal = fRS("Harga").Value
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk = """
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "','" & fTarifTotal & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPenjaminDB & "','" & fJmlTanggunganRSDB & "','" & fJmlPembebasanDB & "',null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fTarifTotal & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPenjaminDB & "',JmlTanggunganRS='" & fJmlTanggunganRSDB & "',JmlPembebasan='" & fJmlPembebasanDB & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk = """
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    'end Tarif Total
    
    'begin Tarif Cito
    If fStatusCito = "1" Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fKdKomponenTarifCito = "07"
        Else
            fKdKomponenTarifCito = fRS("KdKomponenTarifCito").Value
        End If
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlTanggunganRSDB
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "','" & fTarifCito & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fTarifCito & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk = """
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
    End
    End If
    'end Tarif Cito
End Function
'Konversi dari SP: Add_TempHargaKomponenForPenunjangM
Public Function f_AddTempHargaKomponenForPenunjangM(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Currency, fJmlPelayanan As Integer, fStatusCito As String, fKdLaboratory As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fIdPegawai1 As String
    Dim fIdPegawai2 As String
    Dim fIdPegawai3 As String
    Dim fKdJenisPegawai1 As String
    Dim fKdJenisPegawai2 As String
    Dim fKdJenisPegawai3 As String
    Dim fJmlPembebasanPerKomp As Currency
    Dim fJmlHutangPerKomp As Currency
    Dim fJmlTanggunganPerKomp As Currency
    Dim fTarifKelasPenjaminDB As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalTarifPenjamin As Currency
    Dim fKdRuanganAsal As String
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    
    Set fRS = Nothing
    fQuery = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPegawai1 = fRS("IdPegawai").Value
    fIdPegawai2 = fRS("IdPegawai2").Value
    fIdPegawai3 = fRS("IdPegawai3").Value
    fTarifKelasPenjaminDB = fRS("TarifKelasPenjamin").Value
    fJmlHutangPenjaminDB = fRS("JmlHutangPenjamin").Value
    fJmlTanggunganRSDB = fRS("JmlTanggunganRS").Value
    fJmlPembebasanDB = fRS("JmlPembebasan").Value
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai1 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai2 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai2 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai3 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai3 = fRS("KdJenisPegawai").Value
    fTotalTarifPenjamin = fTarifKelasPenjaminDB + fTarifCito
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    If fKdJenisPegawai1 = "001" Then
        fIdDokter = fIdPegawai
    Else
        fIdDokter = Null
    End If
    If fKdLaboratory = "" Then
        Set fRS = Nothing
        fQuery = "select KdPelayananRS from ConvertPelayananToJasaDokter where KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
        Else
            If (fIdDokter = "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14')"
            End If
            If (fIdPegawai2 = "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14')"
            End If
            If (fIdPegawai2 <> "") And (fIdPegawai3 = "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen<>'14'"
            End If
            If (fIdPegawai2 <> "") And (fIdPegawai3 <> "") And (fIdDokter <> "") Then
                fQuery = "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas='" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'"
            End If
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
        While fRS.EOF = False
            fKdKomponen = fRS("KdKomponen").Value
            Set fRS2 = Nothing
            fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fKdKomponen & "') as Harga"
            Call msubRecFO(fRS2, fQuery2)
            fHarga = fRS2("Harga").Value
            If fHarga = "" Then
                fHarga = 0
            End If
            fJmlPembebasanPerKomp = 0
            If fTarifKelasPenjaminDB = 0 Then
                fJmlHutangPerKomp = 0
                fJmlTanggunganPerKomp = 0
            Else
                fJmlHutangPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
                fJmlTanggunganPerKomp = (fHarga / fTotalTarifPenjamin) * fJmlTanggunganRSDB
            End If
            Set fRS2 = Nothing
            fQuery2 = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
            Call msubRecFO(fRS2, fQuery2)
            If fRS2.EOF = True Then
                If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
                End If
                If fKdKomponen = "04" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai2 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
                End If
                If fKdKomponen = "14" Then
                    fQuery2 = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponen & "','" & fKdJenisTarif & "','" & fHarga & "','" & fJmlPelayanan & "', null,'" & fIdPegawai3 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
                End If
            Else
                If fKdKomponen <> "04" And fKdKomponen <> "14" Then
                   fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
                End If
                If fKdKomponen = "04" Then
                   fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai2 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
                End If
                If fKdKomponen = "14" Then
                   fQuery2 = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fHarga & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai3 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and NoStruk = """
                End If
            End
            End If
            Set fRS2 = Nothing
            Call msubRecFO(fRS2, fQuery2)
            fRS.MoveNext
        End
        Wend
    Else
        Set fRS = Nothing
        fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "','" & fIdPegawai1 & "','" & fIdPegawai2 & "','" & fIdPegawai3 & "', 'T') as Harga"
        Call msubRecFO(fRS, fQuery)
        fTarifTotal = fRS("Harga").Value
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='16' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','16','" & fKdJenisTarif & "','" & fTarifTotal & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPenjaminDB & "','" & fJmlTanggunganRSDB & "','" & fJmlPembebasanDB & "',null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fTarifTotal & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPenjaminDB & "',JmlTanggunganRS='" & fJmlTanggunganRSDB & "',JmlPembebasan='" & fJmlPembebasanDB & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='16' and NoStruk = """
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
    End
    End If
    '--begin Tarif Total
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalTM from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fKdKomponenTarifTotal = "12"
    Else
        fKdKomponenTarifTotal = fRS("KdKomponenTarifTotalTM").Value
    End If
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTM('" & fNoPendaftaran & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','" & fStatusCito & "','" & fIdPegawai1 & "','" & fIdPegawai2 & "','" & fIdPegawai3 & "', 'T') as Harga"
    Call msubRecFO(fRS, fQuery)
    fTarifTotal = fRS("Harga").Value
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & fTglPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk = """
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifTotal & "','" & fKdJenisTarif & "','" & fTarifTotal & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPenjaminDB & "','" & fJmlTanggunganRSDB & "','" & fJmlPembebasanDB & "',null)"
    Else
        fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fTarifTotal & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPenjaminDB & "',JmlTanggunganRS='" & fJmlTanggunganRSDB & "',JmlPembebasan='" & fJmlPembebasanDB & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifTotal & "' and NoStruk = """
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
    'end Tarif Total
    
    'begin Tarif Cito
    If fStatusCito = "1" Then
        Set fRS = Nothing
        fQuery = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fKdKomponenTarifCito = "07"
        Else
            fKdKomponenTarifCito = fRS("KdKomponenTarifCito").Value
        End If
        fJmlPembebasanPerKomp = 0
        If fTarifKelasPenjaminDB = 0 Then
            fJmlHutangPerKomp = 0
            fJmlTanggunganPerKomp = 0
        Else
            fJmlHutangPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKomp = (fTarifCito / fTotalTarifPenjamin) * fJmlTanggunganRSDB
        End If
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            fQuery = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdKomponenTarifCito & "','" & fKdJenisTarif & "','" & fTarifCito & "','" & fJmlPelayanan & "', null,'" & fIdPegawai1 & "','" & fJmlHutangPerKomp & "','" & fJmlTanggunganPerKomp & "','" & fJmlPembebasanPerKomp & "',null)"
        Else
            fQuery = "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga='" & fTarifCito & "',JmlPelayanan='" & fJmlPelayanan & "',IdPegawai='" & fIdPegawai1 & "',JmlHutangPenjamin='" & fJmlHutangPerKomp & "',JmlTanggunganRS='" & fJmlTanggunganPerKomp & "',JmlPembebasan='" & fJmlPembebasanPerKomp & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "' and NoStruk = """
        End If
        Set fRS = Nothing
        Call msubRecFO(fRS, fQuery)
    End If
    'end Tarif Cito
End Function
'Konversi dari SP: AM_DataPelayananApotikPH
Public Function f_AMDataPelayananApotikPH(fNoStruk As String, fTglStruk As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fHarga As Currency, fJmlService As Integer, fJmlBarang As Double, fStatus As String)
    'fStatus=A=Add; M=Min
    Dim fTotalHarga As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdPelayananRS As String
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fTotalBiaya As Currency
    Dim fTarifService As Currency
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fDiscount As Currency
    Dim fHargaAkhir As Currency
    Dim fHargaSatuan As Currency
    
    
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    If fIdPenjamin = "" Or fKdKelompokPasien = "" Then
        fIdPenjamin = "2222222222"
        fKdKelompokPasien = "01"
    End If
    Set fRS = Nothing
    fQuery = "select TarifService,JmlHutangPenjamin,JmlTanggunganRS,Discount,HargaSatuan from ApotikJual where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and SatuanJml='" & fSatuanJml & "' and KdAsal='" & fKdAsal & "'"
    Call msubRecFO(fRS, fQuery)
    fTarifService = fRS("TarifService").Value
    fJmlHutangPenjamin = fRS("JmlHutangPenjamin").Value
    fJmlTanggunganRS = fRS("JmlTanggunganRS").Value
    fDiscount = fRS("Discount").Value
    fHargaSatuan = fRS("HargaSatuan").Value
    fHargaAkhir = fHargaSatuan - fDiscount
    fTotalHarga = (fTarifService + fHargaAkhir)
    If fKdKomponen = "10" Then
        fTotalBiaya = (fHarga * fJmlService)
        fJmlHutangPenjaminTotal = fJmlService * ((fHarga / fTotalHarga) * fJmlHutangPenjamin)
        fJmlTanggunganRSTotal = fJmlService * ((fHarga / fTotalHarga) * fJmlTanggunganRS)
    Else
        fTotalBiaya = (fJmlBarang * fHarga)
        fJmlHutangPenjaminTotal = fJmlBarang * ((fHarga / fTotalHarga) * fJmlHutangPenjamin)
        fJmlTanggunganRSTotal = fJmlBarang * ((fHarga / fTotalHarga) * fJmlTanggunganRS)
    End If
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananApotikPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananApotikPH values('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "','" & fJmlBarang & "','" & fTotalBiaya & "','" & fJmlHutangPenjaminTotal & "','" & fJmlTanggunganRSTotal & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananApotikPH set JmlBarang=JmlBarang+'" & fJmlBarang & "',TotalBiaya=TotalBiaya+'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin+'" & fJmlHutangPenjaminTotal & "',TotalTanggunganRS=TotalTanggunganRS+'" & fJmlTanggunganRSTotal & "' where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananApotikPH set JmlBarang=JmlBarang-'" & fJmlBarang & "',TotalBiaya=TotalBiaya-'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin-'" & fJmlHutangPenjaminTotal & "',TotalTanggunganRS=TotalTanggunganRS-'" & fJmlTanggunganRSTotal & "' where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglStruk)=datepart(hh, '" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and day(TglStruk)=day('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and month(TglStruk)=month('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "') and year(TglStruk)=year('" & Format(fTglStruk, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function
'Konversi dari SP: AM_DataPelayananOAPasienPH
Public Function f_AMDataPelayananOAPasienPH(fNoPendaftaran As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fJmlService As Integer, fJmlBarang As Double, fStatus As String)
    'fStatus = A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String
    Dim fKdKelas As String
    Dim fKdPelayananRS As String
    
    Set fRS = Nothing
    fQuery = "select KdPelayananRSOA from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fKdPelayananRS = "000001"
    Else
        fKdPelayananRS = fRS("KdPelayananRSOA").Value
        If fKdPelayananRS = "" Then fKdPelayananRS = "000001"
    End If
    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisKelamin = fRS("KdJenisKelamin").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    If fIdPenjamin = "" Then
        fIdPenjamin = "2222222222"
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,KdKelas from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and SatuanJml='" & fSatuanJml & "' and KdAsal='" & fKdAsal & "'"
    Call msubRecFO(fRS, fQuery)
    fKdSubInstalasi = fRS("KdSubInstalasi").Value
    fKdKelas = fRS("KdKelas").Value
    If fKdKomponen = "10" Then
        fTotalBiaya = fJmlService * fHarga
        fTotalHutangPenjamin = fJmlService * fJmlHutangPenjamin
        fTotalTanggunganRS = fJmlService * fJmlTanggunganRS
        fTotalPembebasan = fJmlService * fJmlPembebasan
    Else
        fTotalBiaya = fJmlBarang * fHarga
        fTotalHutangPenjamin = fJmlBarang * fJmlHutangPenjamin
        fTotalTanggunganRS = fJmlBarang * fJmlTanggunganRS
        fTotalPembebasan = fJmlBarang * fJmlPembebasan
    End If
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananOAPasienPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananOAPasienPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & fKdBarang & "','" & fKdKomponen & "','" & fKdJenisKelamin & "','" & fJmlBarang & "','" & fTotalBiaya & "','" & fTotalHutangPenjamin & "','" & fTotalTanggunganRS & "','" & fTotalPembebasan & "','" & fKdPelayananRS & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananOAPasienPH set JmlBarang=JmlBarang+'" & fJmlBarang & "',TotalBiaya=TotalBiaya+'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin+'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS+'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan+'" & fTotalPembebasan & "'" _
                    & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananOAPasienPH set JmlBarang=JmlBarang-'" & fJmlBarang & "',TotalBiaya=TotalBiaya-'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin-'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS-'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan-'" & fTotalPembebasan & "' " _
                    & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function
'Konversi dari SP: AM_DataPelayananTMPasienPH
Public Function f_AMDataPelayananTMPasienPH(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fKdKelas As String, fStatus As String)
    'fStatus= A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fJmlPelayanan As Integer
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String
    
    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisKelamin = fRS("KdJenisKelamin").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    If fIdPenjamin = "" Then
        fIdPenjamin = "2222222222"
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,StatusAPBD,JmlPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fKdSubInstalasi = fRS("KdSubInstalasi").Value
    fKdAsal = fRS("StatusAPBD").Value
    fJmlPelayanan = fRS("JmlPelayanan").Value
    
    fTotalBiaya = fJmlPelayanan * fHarga
    fTotalHutangPenjamin = fJmlPelayanan * fJmlHutangPenjamin
    fTotalTanggunganRS = fJmlPelayanan * fJmlTanggunganRS
    fTotalPembebasan = fJmlPelayanan * fJmlPembebasan
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananTMPasienPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananTMPasienPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & KdPelayananRS & "','" & fKdKomponen & "','" & fKdJenisKelamin & "','" & fJmlPelayanan & "','" & fTotalBiaya & "','" & fTotalHutangPenjamin & "','" & fTotalTanggunganRS & "','" & TotalPembebasan & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananTMPasienPH set JmlPelayanan=JmlPelayanan+'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya+'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin+'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS+'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan+'" & fTotalPembebasan & "'" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananTMPasienPH set JmlPelayanan=JmlPelayanan-'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya-'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin-'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS-'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan-'" & fTotalPembebasan & "'" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function
'Konversi dari SP: AM_DataPelayananTMPasienDokterPH
Public Function f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String, fKdRuanganAsal As String, fKdKomponen As String, fHarga As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fKdKelas As String, fIdPegawai As String, fStatus As String)
    'fStatus= A:Add; M:Min
    Dim fTotalBiaya As Currency
    Dim fTotalHutangPenjamin As Currency
    Dim fTotalTanggunganRS As Currency
    Dim fTotalPembebasan As Currency
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fJmlPelayanan As Integer
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisKelamin As String
    
    Set fRS = Nothing
    fQuery = "select KdJenisKelamin from V_JenisKelaminPasienTerdaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisKelamin = fRS("KdJenisKelamin").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien,KdDetailJenisJasaPelayanan from V_JenisPasienNPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    If fIdPenjamin = "" Then
        fIdPenjamin = "2222222222"
    End If
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi,StatusAPBD,JmlPelayanan from DetailBiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fKdSubInstalasi = fRS("KdSubInstalasi").Value
    fKdAsal = fRS("StatusAPBD").Value
    fJmlPelayanan = fRS("JmlPelayanan").Value
    
    fTotalBiaya = fJmlPelayanan * fHarga
    fTotalHutangPenjamin = fJmlPelayanan * fJmlHutangPenjamin
    fTotalTanggunganRS = fJmlPelayanan * fJmlTanggunganRS
    fTotalPembebasan = fJmlPelayanan * fJmlPembebasan
    Set fRS = Nothing
    fQuery = "select KdRuangan from DataPelayananTMPasienDokterPH where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery = "insert into DataPelayananTMPasienDokterPH values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdSubInstalasi & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdAsal & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fIdPegawai & "','" & fKdJenisKelamin & "','" & fJmlPelayanan & "','" & fTotalBiaya & "','" & fTotalHutangPenjamin & "','" & fTotalTanggunganRS & "','" & fTotalPembebasan & "')"
    Else
        If UCase(fStatus) = "A" Then
            fQuery = "update DataPelayananTMPasienDokterPH set JmlPelayanan=JmlPelayanan+'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya+'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin+'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS+'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan+'" & fTotalPembebasan & "'" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery = "update DataPelayananTMPasienDokterPH set JmlPelayanan=JmlPelayanan-'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya-'" & fTotalBiaya & "',TotalHutangPenjamin=TotalHutangPenjamin-'" & fTotalHutangPenjamin & "',TotalTanggunganRS=TotalTanggunganRS-'" & fTotalTanggunganRS & "',TotalPembebasan=TotalPembebasan-'" & fTotalPembebasan & "'" _
            & "where (KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "' and KdJenisKelamin='" & fKdJenisKelamin & "' and IdPegawai='" & fIdPegawai & "') and (datepart(hh, TglPelayanan)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglPelayanan)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglPelayanan)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglPelayanan)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS = Nothing
    Call msubRecFO(fRS, fQuery)
End Function
'Konversi dari SP: Add_TempHargaKomponenOAResep
Public Function f_AddTempHargaKomponenOAResep(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fHargaSatuan As Currency, fHargaBeli As Currency, fJmlBarang As Double, fKdJenisObat As String, fJmlService As Integer, fTarifService As Currency, fNoResep As String, fBiayaAdministrasi As Currency, fKdRuanganAsal As String)
    Dim fKdKomponenProfit As String
    Dim fKdKomponenTotal As String
    Dim fKdKomponenHargaNetto As String
    Dim fHargaBersih As Currency
    Dim fKdKomponenTarifService As String
    Dim fKdKomponenAdm As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fTarifServiceResep As Currency
    Dim fJasaRS As Currency
    Dim fJmlPembebasanPerKompP As Currency
    Dim fJmlHutangPerKompP As Currency
    Dim fJmlTanggunganPerKompP As Currency
    Dim fJmlPembebasanPerKompHN As Currency
    Dim fJmlHutangPerKompHN As Currency
    Dim fJmlTanggunganPerKompHN As Currency
    Dim fJmlPembebasanPerKompTotal As Currency
    Dim fJmlHutangPerKompTotal As Currency
    Dim fJmlTanggunganPerKompTotal As Currency
    Dim fJmlPembebasanPerKompAdm As Currency
    Dim fJmlHutangPerKompAdm As Currency
    Dim fJmlTanggunganPerKompAdm As Currency
    Dim fJmlPembebasanPerKompService As Currency
    Dim fJmlHutangPerKompService As Currency
    Dim fJmlTanggunganPerKompService As Currency
    Dim fJmlPembebasanPerKompRS As Currency
    Dim fJmlHutangPerKompRS As Currency
    Dim fJmlTanggunganPerKompRS As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalHarga As Currency
    
    
    Set fRS = Nothing
    fQuery = "select JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from DetailPemakaianAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fJmlHutangPenjaminDB = fRS("JmlHutangPenjamin").Value
    fJmlTanggunganRSDB = fRS("JmlTanggunganRS").Value
    fJmlPembebasanDB = fRS("JmlPembebasan").Value
    Set fRS = Nothing
    fQuery = "select KdKelompokPasien,IdPenjamin from V_KelasTanggunganPenjamin where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    If fIdPenjamin = "" Then fIdPenjamin = "2222222222"
    fHargaBersih = fHargaSatuan - fHargaBeli
    fTotalHarga = fHargaSatuan + fTarifService + fBiayaAdministrasi
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalOA,KdKomponenProfit,KdKomponenHargaNetto,KdKomponenTarifServisResep,KdKomponenAdm from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    fKdKomponenTotal = fRS("KdKomponenTarifTotalOA").Value
    fKdKomponenProfit = fRS("KdKomponenProfit").Value
    fKdKomponenHargaNetto = fRS("KdKomponenHargaNetto").Value
    fKdKomponenTarifService = fRS("KdKomponenTarifServisResep").Value
    fKdKomponenAdm = fRS("KdKomponenAdm").Value
    If fKdKomponenProfit = "" Then fKdKomponenProfit = "13"
    If fKdKomponenHargaNetto = "" Then fKdKomponenHargaNetto = "09"
    If fKdKomponenTotal = "" Then fKdKomponenTotal = "06"
    If fKdKomponenTarifService = "" Then fKdKomponenTarifService = "10"
    If fKdKomponenAdm = "" Then fKdKomponenAdm = "22"
   'begin Tarif Total
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "' and NoStruk = """
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlPembebasanDB
        Else
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fHargaSatuan
        End If
        fJmlHutangPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTotal & "','" & fHargaSatuan & "','" & fJmlBarang & "',null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompTotal & "','" & fJmlTanggunganPerKompTotal & "','" & fJmlPembebasanPerKompTotal & "',null)"
    Else
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlPembebasanDB
        Else
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fHargaSatuan
        End If
        fJmlHutangPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompTotal & "',JmlTanggunganRS='" & fJmlTanggunganPerKompTotal & "',JmlPembebasan='" & fJmlPembebasanPerKompTotal & "',HargaSatuan='" & fHargaSatuan & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "',NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "' and NoStruk = """
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
   'end Tarif Total
   'begin Harga Netto
    Set fRS = Nothing
    fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "' and NoStruk = """
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
    
        If fJmlPembebasanDB <= fHargaBeli Then
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fJmlPembebasanDB
        Else
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fHargaBeli
        End If
        fJmlHutangPerKompHN = (fHargaBeli / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompHN = (fHargaBeli / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenHargaNetto & "','" & fHargaBeli & "','" & fJmlBarang & "',null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompHN & "','" & fJmlTanggunganPerKompHN & "','" & fJmlPembebasanPerKompHN & "',null)"
    End
    Else
    
        If fJmlPembebasanDB <= fHargaSatuan Then
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fJmlPembebasanDB
        Else
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fHargaBeli
        End If
        fJmlHutangPerKompHN = (fHargaBeli / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompHN = (fHargaBeli / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompHN & "',JmlTanggunganRS='" & fJmlTanggunganPerKompHN & "',JmlPembebasan='" & fJmlPembebasanPerKompHN & "',HargaSatuan='" & fHargaBeli & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "',NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "' and NoStruk = """
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
   'end Harga Netto
   'begin Profit atau Keuntungan
    If fHargaBersih <> 0 Then
    
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
        
            If fJmlPembebasanDB > fHargaBeli Then
                fJmlPembebasanPerKompP = (fHargaBersih / fTotalHarga) * (fJmlPembebasanDB - fHargaBeli)
            Else
                fJmlPembebasanPerKompP = 0
             End If
            fJmlHutangPerKompP = (fHargaBersih / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompP = (fHargaBersih / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = " insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenProfit & "','" & fHargaBersih & "','" & fJmlBarang & "',null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompP & "','" & fJmlTanggunganPerKompP & "','" & fJmlPembebasanPerKompP & "',null)"
        End
        Else
        
            If fJmlPembebasanDB > fHargaBeli Then
                fJmlPembebasanPerKompP = (fHargaBersih / fTotalHarga) * (fJmlPembebasanDB - fHargaBeli)
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (fHargaBersih / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompP = (fHargaBersih / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompP & "',JmlTanggunganRS='" & fJmlTanggunganPerKompP & "',JmlPembebasan='" & fJmlPembebasanPerKompP & "',HargaSatuan='" & fHargaBersih & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "',NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "' and NoStruk = """
        End
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End
    End If
   'end Profit atau Keuntungan
   'begin Tarif Service Resep
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat='" & fKdJenisObat & "' and KdKomponen='" & fKdKomponenTarifService & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fTarifServiceResep = 0
    Else
        fTarifServiceResep = fRS("TarifService").Value
    End If
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat='" & fKdJenisObat & "' and KdKomponen='01' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fJasaRS = 0
    Else
        fJasaRS = fRS("TarifService").Value
    End If
    If (fTarifServiceResep = 0 And fJasaRS = 0) And fTarifService <> 0 Then
        fTarifServiceResep = fTarifService
    End If
    If fTarifServiceResep <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > fHargaSatuan Then
                fJmlPembebasanPerKompService = (fTarifServiceResep / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan)
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTarifService & "','" & fTarifServiceResep & "','" & fJmlService & "',null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompService & "','" & fJmlTanggunganPerKompService & "','" & fJmlPembebasanPerKompService & "',null)"
        Else
        
            If fJmlPembebasanDB > fHargaSatuan Then
                fJmlPembebasanPerKompService = (fTarifServiceResep / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan)
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompService & "',JmlTanggunganRS='" & fJmlTanggunganPerKompService & "',JmlPembebasan='" & fJmlPembebasanPerKompService & "',HargaSatuan='" & fTarifServiceResep & "',JmlBarang='" & fJmlService & "',KdJenisObat='" & fKdJenisObat & "',NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "' and NoStruk = """
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    If fJasaRS <> 0 And fJasaRS <> "" Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (fJasaRS / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan - fTarifServiceResep)
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (fJasaRS / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompRS = (fJasaRS / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','01','" & fJasaRS & "','" & fJmlService & "',null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompRS & "','" & fJmlTanggunganPerKompRS & "','" & fJmlPembebasanPerKompRS & "',null)"
        Else
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (fJasaRS / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan - fTarifServiceResep)
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (fJasaRS / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompRS = (fJasaRS / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompRS & "',JmlTanggunganRS='" & fJmlTanggunganPerKompRS & "',JmlPembebasan='" & fJmlPembebasanPerKompRS & "',HargaSatuan='" & fJasaRS & "',JmlBarang='" & fJmlService & "',KdJenisObat='" & fKdJenisObat & "',NoResep='" & fNoResep & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01' and NoStruk = """
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
   'end Tarif Service Resep
   'begin Biaya Administrasi
    If fBiayaAdministrasi <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoPendaftaran from TempHargaKomponenObatAlkes where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "' and NoStruk = """
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan - fTarifServiceResep - fJasaRS)
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenObatAlkes values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenAdm & "','" & fBiayaAdministrasi & "',1,null,'" & fKdJenisObat & "','" & fNoResep & "','" & fJmlHutangPerKompAdm & "','" & fJmlTanggunganPerKompAdm & "','" & fJmlPembebasanPerKompAdm & "',null)"
        Else
        
            If fJmlPembebasanDB > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * (fJmlPembebasanDB - fHargaSatuan - fTarifServiceResep - fJasaRS)
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenObatAlkes set JmlHutangPenjamin='" & fJmlHutangPerKompAdm & "',JmlTanggunganRS='" & fJmlTanggunganPerKompAdm & "',JmlPembebasan='" & fJmlPembebasanPerKompAdm & "',HargaSatuan='" & fBiayaAdministrasi & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:dd") & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "' and NoStruk = """
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
   'end Biaya Administrasi
End Function
'Konversi dari SP: Add_TempHargaKomponenApotik
Public Function f_AddTempHargaKomponenApotik(fNoStruk As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fHargaSatuan As Currency, fHargaBeli As Currency, fJmlBarang As Double, fKdJenisObat As String, fJmlService As Integer, fTarifService As Currency, fBiayaAdministrasi As Currency)
    Dim fKdKomponenProfit As String
    Dim fKdKomponenTotal As String
    Dim fKdKomponenHargaNetto As String
    Dim fHargaBersih As Currency
    Dim fKdKomponenTarifService As String
    Dim fKdRuanganAsal As String
    Dim fTglStruk As Date
    Dim fKdKomponenAdm As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fTarifServiceResep As Currency
    Dim fJasaRS As Currency
    Dim fDiscount As Currency
    Dim fJmlPembebasanPerKompP As Currency
    Dim fJmlHutangPerKompP As Currency
    Dim fJmlTanggunganPerKompP As Currency
    Dim fJmlPembebasanPerKompHN As Currency
    Dim fJmlHutangPerKompHN As Currency
    Dim fJmlTanggunganPerKompHN As Currency
    Dim fJmlPembebasanPerKompTotal As Currency
    Dim fJmlHutangPerKompTotal As Currency
    Dim fJmlTanggunganPerKompTotal As Currency
    Dim fJmlPembebasanPerKompAdm As Currency
    Dim fJmlHutangPerKompAdm As Currency
    Dim fJmlTanggunganPerKompAdm As Currency
    Dim fJmlPembebasanPerKompService As Currency
    Dim fJmlHutangPerKompService As Currency
    Dim fJmlTanggunganPerKompService As Currency
    Dim fJmlPembebasanPerKompRS As Currency
    Dim fJmlHutangPerKompRS As Currency
    Dim fJmlTanggunganPerKompRS As Currency
    Dim fJmlHutangPenjaminDB As Currency
    Dim fJmlTanggunganRSDB As Currency
    Dim fJmlPembebasanDB As Currency
    Dim fTotalPembebasan As Currency
    Dim fTotalHarga As Currency
    
    
    Set fRS = Nothing
    fQuery = "select TglStruk,KdRuanganAsal,KdKelompokPasien,IdPenjamin from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    fTglStruk = fRS("TglStruk").Value
    fKdKelompokPasien = fRS("KdKelompokPasien").Value
    fIdPenjamin = fRS("IdPenjamin").Value
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select Discount,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from ApotikJual where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "'"
    Call msubRecFO(fRS, fQuery)
    fDiscount = fRS("Discount").Value
    fJmlHutangPenjaminDB = fRS("JmlHutangPenjaminDB").Value
    fJmlTanggunganRSDB = fRS("JmlTanggunganRSDB").Value
    fJmlPembebasanDB = fRS("JmlPembebasanDB").Value
    
    fHargaBersih = fHargaSatuan - fHargaBeli
    fTotalPembebasan = fJmlPembebasanDB + fDiscount
    fTotalHarga = fHargaSatuan + fTarifService + fBiayaAdministrasi
    Set fRS = Nothing
    fQuery = "select KdKomponenTarifTotalOA,KdKomponenProfit,KdKomponenHargaNetto,KdKomponenTarifServisResep,KdKomponenAdm from MasterDataPendukung"
    Call msubRecFO(fRS, fQuery)
    fKdKomponenTotal = fRS("KdKomponenTarifTotalOA").Value
    fKdKomponenProfit = fRS("KdKomponenProfit").Value
    fKdKomponenHargaNetto = fRS("KdKomponenHargaNetto").Value
    fKdKomponenTarifService = fRS("KdKomponenTarifServisResep").Value
    fKdKomponenAdm = fRS("KdKomponenAdm").Value
    If fKdKomponenProfit = "" Then fKdKomponenProfit = "13"
    If fKdKomponenHargaNetto = "" Then fKdKomponenHargaNetto = "09"
    If fKdKomponenTotal = "" Then fKdKomponenTotal = "06"
    If fKdKomponenTarifService = "" Then fKdKomponenTarifService = "10"
    If fKdKomponenAdm = "" Then fKdKomponenAdm = "22"
   'begin Tarif Total
    Set fRS = Nothing
    fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
    
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fTotalPembebasan
        Else
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fHargaSatuan
        End If
        fJmlHutangPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTotal & "','" & fJmlBarang & "','" & fHargaSatuan & "','" & fKdJenisObat & "','" & fJmlHutangPerKompTotal & "','" & fJmlTanggunganPerKompTotal & "','" & fJmlPembebasanPerKompTotal & "',null)"
    End
    Else
    
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fTotalPembebasan
        Else
            fJmlPembebasanPerKompTotal = (fHargaSatuan / fTotalHarga) * fHargaSatuan
        End If
        fJmlHutangPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompTotal = (fHargaSatuan / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompTotal & "',JmlTanggunganRS='" & fJmlTanggunganPerKompTotal & "',JmlPembebasan='" & fJmlPembebasanPerKompTotal & "',HargaSatuan='" & fHargaSatuan & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTotal & "'"
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
   'end Tarif Total
   'begin Harga Netto
    Set fRS = Nothing
    fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
    
        If fTotalPembebasan <= fHargaBeli Then
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fTotalPembebasan
        Else
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fHargaBeli
        End If
        fJmlHutangPerKompHN = (fHargaBeli / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompHN = (fHargaBeli / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenHargaNetto & "','" & fJmlBarang & "','" & fHargaBeli & "','" & fKdJenisObat & "','" & fJmlHutangPerKompHN & "','" & fJmlTanggunganPerKompHN & "','" & fJmlPembebasanPerKompHN & "',null)"
    End
    Else
    
        If fTotalPembebasan <= fHargaSatuan Then
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fTotalPembebasan
        Else
            fJmlPembebasanPerKompHN = (fHargaBeli / fTotalHarga) * fHargaBeli
        End If
        fJmlHutangPerKompHN = (fHargaBeli / fTotalHarga) * fJmlHutangPenjaminDB
        fJmlTanggunganPerKompHN = (fHargaBeli / fTotalHarga) * fJmlTanggunganRSDB
        fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompHN & "',JmlTanggunganRS='" & fJmlTanggunganPerKompHN & "',JmlPembebasan='" & fJmlPembebasanPerKompHN & "',HargaSatuan='" & fHargaBeli & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenHargaNetto & "'"
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
   'end Harga Netto
   'begin Profit atau Keuntungan
    If fHargaBersih <> 0 Then
    
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
        
            If fTotalPembebasan > fHargaBeli Then
                fJmlPembebasanPerKompP = (fHargaBersih / fTotalHarga) * (fTotalPembebasan - fHargaBeli)
            Else
                fJmlPembebasanPerKompP = 0
             End If
            fJmlHutangPerKompP = (fHargaBersih / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompP = (fHargaBersih / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = " insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenProfit & "','" & fJmlBarang & "','" & fHargaBersih & "','" & fKdJenisObat & "','" & fJmlHutangPerKompP & "','" & fJmlTanggunganPerKompP & "','" & fJmlPembebasanPerKompP & "',null)"
        End
        Else
        
            If fTotalPembebasan > fHargaBeli Then
                fJmlPembebasanPerKompP = (fHargaBersih / fTotalHarga) * (fTotalPembebasan - fHargaBeli)
            Else
                fJmlPembebasanPerKompP = 0
            End If
            fJmlHutangPerKompP = (fHargaBersih / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompP = (fHargaBersih / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompP & "',JmlTanggunganRS='" & fJmlTanggunganPerKompP & "',JmlPembebasan='" & fJmlPembebasanPerKompP & "',HargaSatuan='" & fHargaBersih & "',JmlBarang='" & fJmlBarang & "',KdJenisObat='" & fKdJenisObat & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenProfit & "'"
        End
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End
    End If
   'end Profit atau Keuntungan
   'begin Tarif Service Resep
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat='" & fKdJenisObat & "' and KdKomponen='" & fKdKomponenTarifService & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fTarifServiceResep = 0
    Else
        fTarifServiceResep = fRS("TarifService").Value
    End If
    Set fRS = Nothing
    fQuery = "select TarifService from DetailTarifJenisObat where KdJenisObat='" & fKdJenisObat & "' and KdKomponen='01' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "'"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fJasaRS = 0
    Else
        fJasaRS = fRS("TarifService").Value
    End If
    If (fTarifServiceResep = 0 And fJasaRS = 0) And fTarifService <> 0 Then
        fTarifServiceResep = fTarifService
    End If
    If fTarifServiceResep <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > fHargaSatuan Then
                fJmlPembebasanPerKompService = (fTarifServiceResep / fTotalHarga) * (fTotalPembebasan - fHargaSatuan)
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenTarifService & "','" & fJmlService & "','" & fTarifServiceResep & "','" & fKdJenisObat & "','" & fJmlHutangPerKompService & "','" & fJmlTanggunganPerKompService & "','" & fJmlPembebasanPerKompService & "',null)"
        Else
        
            If fTotalPembebasan > fHargaSatuan Then
                fJmlPembebasanPerKompService = (fTarifServiceResep / fTotalHarga) * (fTotalPembebasan - fHargaSatuan)
            Else
                fJmlPembebasanPerKompService = 0
            End If
            fJmlHutangPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompService = (fTarifServiceResep / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompService & "',JmlTanggunganRS='" & fJmlTanggunganPerKompService & "',JmlPembebasan='" & fJmlPembebasanPerKompService & "',HargaSatuan='" & fTarifServiceResep & "',JmlBarang='" & fJmlService & "',KdJenisObat='" & fKdJenisObat & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenTarifService & "'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
    If fJasaRS <> 0 And fJasaRS <> "" Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (fJasaRS / fTotalHarga) * (fTotalPembebasan - fHargaSatuan - fTarifServiceResep)
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (fJasaRS / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompRS = (fJasaRS / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','01','" & fJmlService & "','" & fJasaRS & "','" & fKdJenisObat & "','" & fJmlHutangPerKompRS & "','" & fJmlTanggunganPerKompRS & "','" & fJmlPembebasanPerKompRS & "',null)"
        Else
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep) Then
                fJmlPembebasanPerKompRS = (fJasaRS / fTotalHarga) * (fTotalPembebasan - fHargaSatuan - fTarifServiceResep)
            Else
                fJmlPembebasanPerKompRS = 0
            End If
            fJmlHutangPerKompRS = (fJasaRS / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompRS = (fJasaRS / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompRS & "',JmlTanggunganRS='" & fJmlTanggunganPerKompRS & "',JmlPembebasan='" & fJmlPembebasanPerKompRS & "',HargaSatuan='" & fJasaRS & "',JmlBarang='" & fJmlService & "',KdJenisObat='" & fKdJenisObat & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='01'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
   'end Tarif Service Resep
    'begin Biaya Administrasi
    If fBiayaAdministrasi <> 0 Then
        Set fRS = Nothing
        fQuery = "select NoStruk from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "'"
        Call msubRecFO(fRS, fQuery)
        If fRS.EOF = True Then
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * (fTotalPembebasan - fHargaSatuan - fTarifServiceResep - fJasaRS)
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "insert into TempHargaKomponenApotik values('" & fNoStruk & "','" & fKdRuangan & "','" & fKdBarang & "','" & fKdAsal & "','" & fSatuanJml & "','" & fKdKomponenAdm & "',1,'" & fBiayaAdministrasi & "','" & fKdJenisObat & "','" & fJmlHutangPerKompAdm & "','" & fJmlTanggunganPerKompAdm & "','" & fJmlPembebasanPerKompAdm & "',null)"
        Else
            If fTotalPembebasan > (fHargaSatuan + fTarifServiceResep + fJasaRS) Then
                fJmlPembebasanPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * (fTotalPembebasan - fHargaSatuan - fTarifServiceResep - fJasaRS)
            Else
                fJmlPembebasanPerKompAdm = 0
            End If
            fJmlHutangPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlHutangPenjaminDB
            fJmlTanggunganPerKompAdm = (fBiayaAdministrasi / fTotalHarga) * fJmlTanggunganRSDB
            fQuery2 = "update TempHargaKomponenApotik set JmlHutangPenjamin='" & fJmlHutangPerKompAdm & "',JmlTanggunganRS='" & fJmlTanggunganPerKompAdm & "',JmlPembebasan='" & fJmlPembebasanPerKompAdm & "',HargaSatuan='" & fBiayaAdministrasi & "' where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and KdKomponen='" & fKdKomponenAdm & "'"
        End If
        Set fRS2 = Nothing
        Call msubRecFO(fRS2, fQuery2)
    End If
   'end Biaya Administrasi
End Function
'Konversi dari SP: Add_TempHargaKomponenIBS
Public Function f_AddTempHargaKomponenIBS(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Integer, fJmlPelayanan As Integer, fStatusCito As String, fIdPegawai As String, fIdPegawaiAnastesi As String, fIdPegawai2 As String)
    'fIdPegawai= IdDokter; fIdPegawaiAnastesi= IdDokterAnastesi; fIdPegawai2= IdDokterPendamping/Pembantu
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fTotalTarif As Currency
    Dim fKdKomponenTarifTotal As String
    Dim fKdKomponenTarifCito As String
    Dim fTarifTotal As Currency
    Dim fKdJenisPegawai As String
    Dim fIdDokter As String
    Dim fKdDetailJenisJasaPelayanan As String
    Dim fKdJenisPelayanan As String
    Dim fJasaDokterPendamping As Currency
    Dim fJmlDokter As Integer
    Dim fHargaJPO As Currency
    Dim fHargaJPA As Currency
    Dim fHargaJPP As Currency
    Dim fHargaJPOAkhir As Currency
    Dim fKdPelayananRSL As String
    Dim fHargaJS As Currency
    Dim fHargaJPOTemp As Currency
    Dim fTotalTarifCito As Currency
    Dim fKdJenisPegawai2 As String
    
    
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai1 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai1 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai2 & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPegawai2 = fRS("KdJenisPegawai").Value
    Set fRS = Nothing
    fQuery = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS='" & fKdPelayananRS & "'"
    Call msubRecFO(fRS, fQuery)
    fKdJenisPelayanan = fRS("KdJnsPelayanan").Value
    Set fRS = Nothing
    fQuery = "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'"
    Call msubRecFO(fRS, fQuery)
    fKdDetailJenisJasaPelayanan = fRS("KdDetailJenisJasaPelayanan").Value
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','02') as Harga"
    Call msubRecFO(fRS, fQuery)
    fHarga = fRS("Harga").Value
    Set fRS = Nothing
    fQuery = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRS & "','" & fKdKelas & "','" & fKdJenisTarif & "','01') as HargaJS"
    Call msubRecFO(fRS, fQuery)
    fHargaJS = fRS("HargaJS").Value
    Set fRS = Nothing
    fQuery = "select count(IdDokter) as JmlDokter from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter='" & fIdPegawai & "'"
    Call msubRecFO(fRS, fQuery)
    fJmlDokter = fRS("JmlDokter").Value
    If fJmlDokter = 0 Then
        fHargaJPOAkhir = fHarga
        fHargaJPA = (40 * fHargaJPOAkhir) / 100
        fHargaJPP = (14 * fHargaJPOAkhir) / 100
        fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
    
        If fKdJenisPegawai = "001" Then
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02','" & fIdPegawai & "','" & fHargaJPOAkhir & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "','" & fHargaJPP & "')"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
            
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "','" & fHargaJPA & "')"
                Call msubRecFO(fRS2, fQuery2)
            End If
        Else
        
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "','" & fHargaJPP & "')"
            Call msubRecFO(fRS2, fQuery2)
        End If
        If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "','" & fJasaDokterPendamping & "')"
            Call msubRecFO(fRS2, fQuery2)
        End If
    End If
    If fJmlDokter = 1 Then
    
        Set fRS2 = Nothing
        fQuery2 = "select max(Harga) as HargaJPO from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter='" & fIdPegawai & "'"
        Call msubRecFO(fRS2, fQuery2)
        fHargaJPO = fRS2("HargaJPO").Value
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter='" & fIdPegawai & "' and Harga='" & fHargaJPO & "'"
        Call msubRecFO(fRS2, fQuery2)
        fKdPelayananRSL = fRS("KdPelayananRS").Value
        If fHarga >= fHargaJPO Then
        
            fHargaJPOAkhir = fHarga * 1.5
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen in('02','04','05','14')"
            Call msubRecFO(fRS2, fQuery2)
            If fKdJenisPegawai = "001" Then
            
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02','" & fIdPegawai & "','" & fHargaJPOAkhir & "')"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "','" & fHargaJPP & "')"
                Call msubRecFO(fRS2, fQuery2)
                If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "','" & fHargaJPA & "')"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
            
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "','" & fHargaJPP & "')"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "','" & fJasaDokterPendamping & "')"
                    Call msubRecFO(fRS2, fQuery2)
                Else
                
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                    Call msubRecFO(fRS2, fQuery2)
                    Set fRS2 = Nothing
                    fQuery2 = "update TempHargaKomponenIBS set Harga='" & fJasaDokterPendamping & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            End If
        Else
        
            fHargaJPOAkhir = fHargaJPO * 1.5
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            If fKdKomponen = "02" And fKdJenisPegawai2 = "001" Then
                fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            End If
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPOAkhir & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='02'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPA & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='04'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPP & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='05'"
            Call msubRecFO(fRS2, fQuery2)
            If fKdJenisPegawai = "001" Then
            
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02','" & fIdPegawai & "',0)"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "',0)"
                Call msubRecFO(fRS2, fQuery2)
                If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "',0)"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "',0)"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS = Nothing
                fQuery = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS, fQuery)
                If fRS.EOF = True Then
                    Set fRS2 = Nothing
                    fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                    Call msubRecFO(fRS2, fQuery2)
                End If
            Else
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                Call msubRecFO(fRS2, fQuery2)
                Set fRS2 = Nothing
                fQuery2 = "update TempHargaKomponenIBS set Harga='" & fJasaDokterPendamping & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
            End If
        End If
    End If
    If fJmlDokter > 1 Then
    
        Set fRS2 = Nothing
        fQuery2 = "select max(Harga) as HargaJPOTemp from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter='" & fIdPegawai & "'"
        Call msubRecFO(fRS2, fQuery2)
        fHargaJPOTemp = fRS2("HargaJPOTemp").Value
        Set fRS2 = Nothing
        fQuery2 = "select KdPelayananRS from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdKomponen='02' and IdDokter='" & fIdPegawai & "' and Harga='" & fHargaJPOTemp & "'"
        Call msubRecFO(fRS2, fQuery2)
        fKdPelayananRSL = fRS2("KdPelayananRS").Value
        Set fRS2 = Nothing
        fQuery2 = "select dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "','" & fKdPelayananRSL & "','" & fKdKelas & "','" & fKdJenisTarif & "','02') as HargaJPO"
        Call msubRecFO(fRS2, fQuery2)
        fHargaJPO = fRS2("HargaJPO").Value
        If fHarga >= fHargaJPO Then
        
            fHargaJPOAkhir = fHarga * 2
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdKomponen in('02','04','05','14')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02','" & fIdPegawai & "','" & fHargaJPOAkhir & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "','" & fHargaJPP & "')"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "','" & fHargaJPA & "')"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "','" & fJasaDokterPendamping & "')"
                    Call msubRecFO(fRS, fQuery)
                Else
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenIBS set Harga='" & fJasaDokterPendamping & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS, fQuery)
                End If
            End If
        Else
            fHargaJPOAkhir = fHargaJPO * 2
            fHargaJPA = (40 * fHargaJPOAkhir) / 100
            fHargaJPP = (14 * fHargaJPOAkhir) / 100
            fJasaDokterPendamping = (20 * fHargaJPOAkhir) / 100
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPOAkhir & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='02'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPA & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='04'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "update TempHargaKomponenIBS set Harga='" & fHargaJPP & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='05'"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','01','" & fIdPegawai & "','" & fHargaJS & "')"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','02','" & fIdPegawai & "',0)"
            Call msubRecFO(fRS2, fQuery2)
            Set fRS2 = Nothing
            fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','05','" & fIdPegawai & "',0)"
            Call msubRecFO(fRS2, fQuery2)
            If (fKdJenisPelayanan = "001" Or fKdJenisPelayanan = "002" Or fKdJenisPelayanan = "003" Or fKdJenisPelayanan = "004" Or fKdJenisPelayanan = "005" Or fKdJenisPelayanan = "006" Or fKdJenisPelayanan = "007") And fIdPegawaiAnastesi <> "" Then
                Set fRS2 = Nothing
                fQuery2 = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','04','" & fIdPegawaiAnastesi & "',0)"
                Call msubRecFO(fRS2, fQuery2)
            End If
            If fKdDetailJenisJasaPelayanan = "02" And fKdJenisPegawai2 = "001" Then
                Set fRS2 = Nothing
                fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                Call msubRecFO(fRS2, fQuery2)
                If fRS2.EOF = True Then
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                    Call msubRecFO(fRS, fQuery)
                Else
                    Set fRS = Nothing
                    fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','14','" & fIdPegawai & "',0)"
                    Call msubRecFO(fRS, fQuery)
                    Set fRS = Nothing
                    fQuery = "update TempHargaKomponenIBS set Harga='" & fJasaDokterPendamping & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRSL & "' and KdKomponen='14'"
                    Call msubRecFO(fRS, fQuery)
                End If
            End If
        End If
    End If
   '--Tarif Cito
    If fStatusCito = "1" Then
        If fKdDetailJenisJasaPelayanan = "02" Then
            fTotalTarifCito = (6 * fHargaJPOAkhir) / 100
        Else
            fTotalTarifCito = 25 * (fHargaJPA + fHargaJPOAkhir) / 100
        End If
        Set fRS2 = Nothing
        fQuery2 = "select KdKomponenTarifCito from MasterDataPendukung"
        Call msubRecFO(fRS2, fQuery2)
        fKdKomponenTarifCito = fRS2("KdKomponenTarifCito").Value
        If fKdKomponenTarifCito = "" Then fKdKomponenTarifCito = "07"
        Set fRS2 = Nothing
        fQuery2 = "select NoPendaftaran from TempHargaKomponenIBS where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "'"
        Call msubRecFO(fRS2, fQuery2)
        If fRS2.EOF = True Then
            Set fRS = Nothing
            fQuery = "insert into TempHargaKomponenIBS values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKomponenTarifCito & "','" & fIdPegawai & "','" & fTotalTarifCito & "')"
            Call msubRecFO(fRS, fQuery)
            If fKdDetailJenisJasaPelayanan = "01" Then
                Set fRS = Nothing
                fQuery = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and IdDokter='" & fIdPegawai & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='05'"
                Call msubRecFO(fRS, fQuery)
            End If
       Else
            Set fRS = Nothing
            fQuery = "update TempHargaKomponenIBS set Harga='" & fTotalTarifCito & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponenTarifCito & "'"
            Call msubRecFO(fRS, fQuery)
            If fKdDetailJenisJasaPelayanan = "01" Then
                Set fRS = Nothing
                fQuery = "update TempHargaKomponenIBS set Harga=0 where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='05' and IdDokter='" & fIdPegawai & "'"
                Call msubRecFO(fRS, fQuery)
            End If
       End If
    End If
End Function
'Konversi dari SP: Delete_TempHargaKomponen
Public Function f_DeleteTempHargaKomponen(fNoPendaftaran As String, fKdPelayananRS As String, fTglPelayanan As Date, fKdRuangan As String)
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fIdPegawai As String
    Dim fKdJenisPegawai As String
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    Dim fNoLab_Rad As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    
    Set fRS = Nothing
    fQuery = "select NoLab_Rad from BiayaPelayanan where NoPendaftaran='" & fNoPendaftaran & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and NoStruk = """
    Call msubRecFO(fRS, fQuery)
    fNoLab_Rad = fRS("NoLab_Rad").Value
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    Set fRS = Nothing
    fQuery = "select KdKelas,Harga,KdKomponen,IdPegawai,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdRuangan='" & fKdRuangan & "' and NoStruk = "" and NoClosing <> """
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
    
        fKdKelas = fRS("KdKelas").Value
        fHarga = fRS("Harga").Value
        fKdKomponen = fRS("KdKomponen").Value
        fIdPegawai = fRS("IdPegawai").Value
        Set fRS2 = Nothing
        fQuery2 = "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai & "'"
        Call msubRecFO(fRS2, fQuery2)
        fKdJenisPegawai = fRS2("KdJenisPegawai").Value
        Call f_AMDataPelayananTMPasienPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, "M")
        If fKdJenisPegawai = "001" Then
            Call f_AMDataPelayananTMPasienDokterPH(fNoPendaftaran, fKdPelayananRS, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, fKdKelas, fIdPegawai, "M")
        End If
        fRS.MoveNext
    Wend
    Set fRS = Nothing
End Function
Public Sub Add_HistoryLoginActivity(strNamaObjekDB)
On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAplikasi", adChar, adParamInput, 3, strKdAplikasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("TglActivity", adDate, adParamInput, , Format(Now, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("HostName", adVarChar, adParamInput, 50, strNamaHostLocal)
        .Parameters.Append .CreateParameter("NamaObjekDB", adVarChar, adParamInput, 200, strNamaObjekDB)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_HistoryLoginActivity"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Sub
hell_:
     Call msubPesanError("-Add_HistoryLoginActivity")
End Sub
Public Sub subSp_HistoryLoginAplikasi(strStatus)
On Error GoTo hell_
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAplikasi", adChar, adParamInput, 3, strKdAplikasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("NamaHostAplikasi", adVarChar, adParamInput, 50, strNamaHostLocal)
        .Parameters.Append .CreateParameter("TglLogin", adDate, adParamInput, , Format(dTglLogin, "yyyy/MM/dd HH:mm:ss"))
       
        If strStatus = "A" Then
            .Parameters.Append .CreateParameter("TglLogout", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglLogout", adDate, adParamInput, , Format(dTglLogout, "yyyy/MM/dd HH:mm:ss"))
        End If
'
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, strStatus)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.AU_HistoryLoginAplikasi"
        .CommandType = adCmdStoredProc
        .Execute
    
        If .Parameters("RETURN_VALUE").Value <> 0 Then
            MsgBox "Ada Kesalahan dalam Hapus Rekap Komponen Biaya Pelayanan", vbCritical, "Validasi"
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    
Exit Sub
hell_:
     Call msubPesanError("-AU_HistoryLoginAplikasi")
End Sub


'Konversi dari SP: Delete_TempHargaKomponenApotik
Public Function f_DeleteTempHargaKomponenApotik(fNoStruk As String, fTglStruk As Date, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String)
    Dim fKdKomponen As String
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fJmlBarang As Double
    Dim fJmlService As Integer

    Set fRS = Nothing
    fQuery = "select KdRuanganAsal from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select KdKomponen,HargaSatuan,JmlBarang from TempHargaKomponenApotik where NoStruk='" & fNoStruk & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and NoClosing <> """
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
    
        fKdKomponen = fRS("KdKomponen").Value
        fHargaSatuan = fRS("HargaSatuan").Value
        fJmlBarang = fRS("JmlBarang").Value
        fJmlService = 1
        Call f_AMDataPelayananApotikPH(fNoStruk, fTglStruk, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fHarga, fJmlService, fJmlBarang, "M")
        fRS.MoveNext
    End
    Wend
    Set fRS = Nothing
End Function
'Konversi dari SP: Delete_TempHargaKomponenObatAlkes
Public Function f_DeleteTempHargaKomponenObatAlkes(fNoPendaftaran As String, fKdBarang As String, fTglPelayanan As Date, fKdRuangan As String, fKdAsal As String, fSatuanJml As String)
    Dim fKdKomponen As String
    Dim fKdKelas As String
    Dim fJmlBarang As Double
    Dim fHarga As Currency
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    Dim fNoLab_Rad As String
    Dim fJmlHutangPenjamin As Currency
    Dim fJmlTanggunganRS As Currency
    Dim fJmlPembebasan As Currency
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "',null,'" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    Set fRS = Nothing
    fQuery = "select JmlBarang,HargaSatuan,KdKomponen,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan from TempHargaKomponenObatAlkes where NoPendaftaran=fNoPendaftaran and KdBarang=fKdBarang and TglPelayanan=fTglPelayanan and KdRuangan=fKdRuangan and KdAsal=fKdAsal and SatuanJml=fSatuanJml and NoStruk = "" and NoClosing <> """
    Call msubRecFO(fRS, fQuery)
    While fRS.EOF = False
    
        fJmlBarang = fRS("JmlBarang").Value
        fHarga = fRS("HargaSatuan").Value
        fKdKomponen = fRS("KdKomponen").Value
        fJmlHutangPenjamin = fRS("JmlHutangPenjamin").Value
        fJmlTanggunganRS = fRS("JmlTanggunganRS").Value
        fJmlPembebasan = fRS("JmlPembebasan").Value
        Call f_AMDataPelayananOAPasienPH(fNoPendaftaran, fTglPelayanan, fKdRuangan, fKdRuanganAsal, fKdBarang, fKdAsal, fSatuanJml, fKdKomponen, fHarga, fJmlHutangPenjamin, fJmlTanggunganRS, fJmlPembebasan, CInt(fJmlBarang), fJmlBarang, "M")
        fRS.MoveNext
    End
    Wend
    Set fRS = Nothing
End Function
'Konversi dari SP: AM_RekapitulasiJasaBPApotik
Public Function f_AMRekapitulasiJasaBPApotik(fNoStruk As String, fNoBKM As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fSatuanJml As String, fKdKomponen As String, fJmlBrg As Double, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdRuanganAsal As String
    Dim fKdPelayananRS As String
    Dim fKdDetailJenisJasaPelayanan As String
    
    fKdPelayananRS = "000001"
    fKdDetailJenisJasaPelayanan = "03"
    Set fRS = Nothing
    fQuery = "select KdRuanganAsal from V_StrukPelayananApotik where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    If fKdRuanganAsal = "" Then fKdRuanganAsal = fKdRuangan
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    fTglBKM = fRS("TglBKM").Value
    fKdRuanganKasir = fRS("KdRuangan").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("fKdKelompokPasien").Value
    If (fIdPenjamin = "") Or (fKdKelompokPasien = "") Then
        fIdPenjamin = "2222222222"
        fKdKelompokPasien = "01"
    End If
    fTotalTarif = fJmlBrg * fTarif
    fJmlBayarTotal = fJmlBrg * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlBrg * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlBrg * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlBrg * fJmlPembebasan
    fSisaTagihanTotal = fJmlBrg * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPApotik where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPApotik values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "','" & fJmlBrg & "','" & fTotalTarif & "','" & fJmlBayarTotal & "','" & fJmlHutangPenjaminTotal & "','" & fJmlTanggunganRSTotal & "','" & fJmlPembebasanTotal & "','" & fSisaTagihanTotal & "','" & fKdPelayananRS & "','" & fKdDetailJenisJasaPelayanan & "')"
    Else
    
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPApotik set JmlBarang=JmlBarang+'" & fJmlBrg & "', TotalBiaya=TotalBiaya+'" & fTotalTarif & "', TotalBayar=TotalBayar+'" & fJmlBayarTotal & "', TotalHutangPenjamin=TotalHutangPenjamin+'" & fJmlHutangPenjaminTotal & "', TotalTanggunganRS=TotalTanggunganRS+'" & fJmlTanggunganRSTotal & "', TotalPembebasan=TotalPembebasan+'" & fJmlPembebasanTotal & "', TotalSisaTagihan=TotalSisaTagihan+'" & fSisaTagihanTotal & "' " _
            & " where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPApotik set JmlBarang=JmlBarang-'" & fJmlBrg & "', TotalBiaya=TotalBiaya-'" & fTotalTarif & "', TotalBayar=TotalBayar-'" & fJmlBayarTotal & "', TotalHutangPenjamin=TotalHutangPenjamin-'" & fJmlHutangPenjaminTotal & "', TotalTanggunganRS=TotalTanggunganRS-'" & fJmlTanggunganRSTotal & "', TotalPembebasan=TotalPembebasan-'" & fJmlPembebasanTotal & "', TotalSisaTagihan=TotalSisaTagihan-'" & fSisaTagihanTotal & "' " _
            & " where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function
'Konversi dari SP: AM_RekapitulasiJasaBPOAForRemunerasiFV
Public Function f_AMRekapitulasiJasaBPOAForRemunerasiFV(fNoStruk As String, fNoBKM As String, fNoPendaftaran As String, fKdRuangan As String, fKdBarang As String, fKdAsal As String, fTglPelayanan As Date, fSatuanJml As String, fKdKomponen As String, fJmlBrg As Double, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As String, fStatus As String)
    'fStatus: A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','OA') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    fTglBKM = fRS("TglBKM").Value
    fKdRuanganKasir = fRS("KdRuangan").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("fKdKelompokPasien").Value
    Set fRS = Nothing
    fQuery = "select KdSubInstalasi from PemakaianAlkes where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and SatuanJml='" & fSatuanJml & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fKdSubInstalasi = fRS("KdSubInstalasi").Value
    fTotalTarif = fJmlBrg * fTarif
    fJmlBayarTotal = fJmlBrg * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlBrg * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlBrg * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlBrg * fJmlPembebasan
    fSisaTagihanTotal = fJmlBrg * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPOA4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPOA4Remunerasi values('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdBarang & "','" & fKdAsal & "','" & fKdKomponen & "','000001','" & fJmlBrg & "','" & fTotalTarif & "','" & fJmlBayarTotal & "','" & fJmlHutangPenjaminTotal & "','" & fJmlTanggunganRSTotal & "','" & fJmlPembebasanTotal & "','" & fSisaTagihanTotal & "')"
    Else
    
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPOA4Remunerasi set JmlBarang=JmlBarang+'" & fJmlBrg & "', TotalBiaya=TotalBiaya+'" & fTotalTarif & "', JmlBayar=JmlBayar+'" & fJmlBayarTotal & "', JmlHutangPenjamin=JmlHutangPenjamin+'" & fJmlHutangPenjaminTotal & "', JmlTanggunganRS=JmlTanggunganRS+'" & fJmlTanggunganRSTotal & "', JmlPembebasan=JmlPembebasan+'" & fJmlPembebasanTotal & "', SisaTagihan=SisaTagihan+'" & fSisaTagihanTotal & "' " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPOA4Remunerasi set JmlBarang=JmlBarang-'" & fJmlBrg & "', TotalBiaya=TotalBiaya-'" & fTotalTarif & "', JmlBayar=JmlBayar-'" & fJmlBayarTotal & "', JmlHutangPenjamin=JmlHutangPenjamin-'" & fJmlHutangPenjaminTotal & "', JmlTanggunganRS=JmlTanggunganRS-'" & fJmlTanggunganRSTotal & "', JmlPembebasan=JmlPembebasan-'" & fJmlPembebasanTotal & "', SisaTagihan=SisaTagihan-'" & fSisaTagihanTotal & "' " _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdBarang='" & fKdBarang & "' and KdAsal='" & fKdAsal & "' and KdKomponen='" & fKdKomponen & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function
'Konversi dari SP: AM_RekapitulasiJasaBPTMForRemunerasiFV
Public Function f_AMRekapitulasiJasaBPTMForRemunerasiFV(fNoBKM As String, fNoStruk As String, fNoPendaftaran As String, fKdRuangan As String, fKdPelayananRS As String, fKdKomponen As String, fTglPelayanan As Date, fJmlPelayanan As Integer, fTarif As Currency, fJmlBayar As Currency, fJmlHutangPenjamin As Currency, fJmlTanggunganRS As Currency, fJmlPembebasan As Currency, fSisaTagihan As Currency, fKdDetailJenisJasaPelayanan As String, fKdKelas As String, fNoLab_Rad As String, fStatus As String)
    'fStatus : A=Tambah; M=Minus
    Dim fTglBKM As Date
    Dim fTotalTarif As Currency
    Dim fJmlBayarTotal As Currency
    Dim fJmlHutangPenjaminTotal As Currency
    Dim fJmlTanggunganRSTotal As Currency
    Dim fJmlPembebasanTotal As Currency
    Dim fSisaTagihanTotal As Currency
    Dim fKdRuanganKasir As String
    Dim fKdKelompokPasien As String
    Dim fIdPenjamin As String
    Dim fKdAsal As String
    Dim fKdSubInstalasi As String
    Dim fKdRuanganAsal As String
    Dim fKdInstalasi As String
    
    Set fRS = Nothing
    fQuery = "select dbo.FB_TakeRuanganAsal('" & fNoPendaftaran & "','" & fKdRuangan & "','" & fNoLab_Rad & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal"
    Call msubRecFO(fRS, fQuery)
    fKdRuanganAsal = fRS("KdRuanganAsal").Value
    Set fRS = Nothing
    fQuery = "select TglBKM,KdRuangan from StrukBuktiKasMasuk where NoBKM='" & fNoBKM & "'"
    Call msubRecFO(fRS, fQuery)
    fTglBKM = fRS("TglBKM").Value
    fKdRuanganKasir = fRS("KdRuangan").Value
    Set fRS = Nothing
    fQuery = "select IdPenjamin,KdKelompokPasien from StrukPelayananPasien where NoStruk='" & fNoStruk & "'"
    Call msubRecFO(fRS, fQuery)
    fIdPenjamin = fRS("IdPenjamin").Value
    fKdKelompokPasien = fRS("fKdKelompokPasien").Value
    Set fRS = Nothing
    fQuery = "select StatusAPBD,KdSubInstalasi from BiayaPelayanan where NoStruk='" & fNoStruk & "' and NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(fRS, fQuery)
    fKdSubInstalasi = fRS("KdSubInstalasi").Value
    fKdAsal = fRS("StatusAPBD").Value
    
    fTotalTarif = fJmlPelayanan * fTarif
    fJmlBayarTotal = fJmlPelayanan * fJmlBayar
    fJmlHutangPenjaminTotal = fJmlPelayanan * fJmlHutangPenjamin
    fJmlTanggunganRSTotal = fJmlPelayanan * fJmlTanggunganRS
    fJmlPembebasanTotal = fJmlPelayanan * fJmlPembebasan
    fSisaTagihanTotal = fJmlPelayanan * fSisaTagihan
    Set fRS = Nothing
    fQuery = "select KdRuangan from RekapitulasiJasaBPTM4Remunerasi where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
    Call msubRecFO(fRS, fQuery)
    If fRS.EOF = True Then
        fQuery2 = "insert into RekapitulasiJasaBPTM4Remunerasi values('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "','" & fKdRuanganKasir & "','" & fKdRuangan & "','" & fKdRuanganAsal & "','" & fKdSubInstalasi & "','" & fKdKelompokPasien & "','" & fIdPenjamin & "','" & fKdDetailJenisJasaPelayanan & "','" & fKdKelas & "','" & fKdPelayananRS & "','" & fKdKomponen & "','" & fKdAsal & "','" & fJmlPelayanan & "','" & fTotalTarif & "','" & fJmlBayarTotal & "','" & fJmlHutangPenjaminTotal & "','" & fJmlTanggunganRSTotal & "','" & fJmlPembebasanTotal & "','" & fSisaTagihanTotal & "')"
    Else
    
        If UCase(fStatus) = "A" Then
            fQuery2 = "update RekapitulasiJasaBPTM4Remunerasi set JmlPelayanan=JmlPelayanan+'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya+'" & fTotalTarif & "', JmlBayar=JmlBayar+'" & fJmlBayarTotal & "', JmlHutangPenjamin=JmlHutangPenjamin+'" & fJmlHutangPenjaminTotal & "', JmlTanggunganRS=JmlTanggunganRS+'" & fJmlTanggunganRSTotal & "', JmlPembebasan=JmlPembebasan+'" & fJmlPembebasanTotal & "', SisaTagihan=SisaTagihan+'" & fSisaTagihanTotal & "'" _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        Else
            fQuery2 = "update RekapitulasiJasaBPTM4Remunerasi set JmlPelayanan=JmlPelayanan-'" & fJmlPelayanan & "',TotalBiaya=TotalBiaya-'" & fTotalTarif & "', JmlBayar=JmlBayar-'" & fJmlBayarTotal & "', JmlHutangPenjamin=JmlHutangPenjamin-'" & fJmlHutangPenjaminTotal & "', JmlTanggunganRS=JmlTanggunganRS-'" & fJmlTanggunganRSTotal & "', JmlPembebasan=JmlPembebasan-'" & fJmlPembebasanTotal & "', SisaTagihan=SisaTagihan-'" & fSisaTagihanTotal & "'" _
            & "where (KdRuanganKasir='" & fKdRuanganKasir & "' and KdRuangan='" & fKdRuangan & "' and KdRuanganAsal='" & fKdRuanganAsal & "' and KdSubInstalasi='" & fKdSubInstalasi & "' and KdKelompokPasien='" & fKdKelompokPasien & "' and IdPenjamin='" & fIdPenjamin & "' and KdDetailJenisJasaPelayanan='" & fKdDetailJenisJasaPelayanan & "' and KdKelas='" & fKdKelas & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & fKdKomponen & "' and KdAsal='" & fKdAsal & "') and (datepart(hh, TglBKM)=datepart(hh, '" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and day(TglBKM)=day('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and month(TglBKM)=month('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "') and year(TglBKM)=year('" & Format(fTglBKM, "yyyy/MM/dd HH:mm:ss") & "'))"
        End If
    End If
    Set fRS2 = Nothing
    Call msubRecFO(fRS2, fQuery2)
End Function

'@hendri - 20140606
Public Function sp_PostingHutangPenjaminPasien_AU(f_NoPendaftaran As String, f_status As String) As Boolean
    sp_PostingHutangPenjaminPasien_AU = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, adInteger, Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "PostingHutangPenjaminPasien_AU"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_PostingHutangPenjaminPasien_AU = False
        End If
    End With
    Call deleteADOCommandParameters(dbcmd)
End Function

