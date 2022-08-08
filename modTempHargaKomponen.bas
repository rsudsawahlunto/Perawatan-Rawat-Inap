Attribute VB_Name = "modTempHargaKomponen"
Public rsLooping As New ADODB.recordset

Public Function functAdd_TempHargaKomponen(fNoPendaftaran As String, fKdRuangan As String, fTglPelayanan As Date, _
    fKdPelayananRS As String, fKdKelas As String, fKdJenisTarif As String, fTarifCito As Integer, _
    fJmlPelayanan As Integer, fStatusCito As String, fIdPegawai As String, fKdRuanganAsal As String) As Boolean
On Error GoTo errLoad

Dim tempKdKomponen As String
Dim tempHarga As Currency
Dim tempNoPendaftaran As String
Dim tempTotalTarif As Currency
Dim tempKdKomponenTarifTotal As String
Dim tempKdKomponenTarifCito As String
Dim tempTarifTotal As Currency
Dim tempKdJenisPegawai As String
Dim tempIdDokter As String
Dim tempKdDetailJenisJasaPelayanan As String
Dim tempIdPegawai1 As String
Dim tempIdPegawai2 As String
Dim tempIdPegawai3 As String
Dim tempKdJenisPegawai1 As String
Dim tempKdJenisPegawai2 As String
Dim tempKdJenisPegawai3 As String
Dim tempKdPelayananRSTemp As String
Dim tempJmlPembebasanPerKomp As Currency
Dim tempJmlHutangPerKomp As Currency
Dim tempJmlTanggunganPerKomp As Currency
Dim tempTarifKelasPenjaminDB As Currency
Dim tempJmlHutangPenjaminDB As Currency
Dim tempJmlTanggunganRSDB As Currency
Dim tempJmlPembebasanDB As Currency
Dim tempTotalTarifPenjamin As Currency

    strSQL = "select IdPegawai,IdPegawai2,IdPegawai3,TarifKelasPenjamin,JmlHutangPenjamin,JmlTanggunganRS,JmlPembebasan " & _
        " from DetailBiayaPelayanan " & _
        " where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and KdPelayananRS='" & fKdPelayananRS & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        tempIdPegawai1 = rs("idpegawai")
        tempIdPegawai2 = IIf(IsNull(rs("IdPegawai2")), "", rs("IdPegawai2"))
        tempIdPegawai3 = IIf(IsNull(rs("IdPegawai3")), "", rs("IdPegawai3"))
        tempTarifKelasPenjaminDB = rs("TarifKelasPenjamin")
        tempJmlHutangPenjaminDB = rs("JmlHutangPenjamin")
        tempJmlTanggunganRSDB = rs("JmlTanggunganRS")
        tempJmlPembebasanDB = rs("JmlPembebasan")
    Else
        tempIdPegawai1 = ""
        tempIdPegawai2 = ""
        tempIdPegawai3 = ""
        tempTarifKelasPenjaminDB = ""
        tempJmlHutangPenjaminDB = ""
        tempJmlTanggunganRSDB = ""
        tempJmlPembebasanDB = ""
    End If
    
    Call msubRecFO(rs, "select KdJenisPegawai from DataPegawai where IdPegawai='" & tempIdPegawai1 & "'")
    If rs.EOF = False Then tempKdJenisPegawai1 = rs("KdJenisPegawai")
    
    Call msubRecFO(rs, "select KdJenisPegawai from DataPegawai where IdPegawai='" & tempIdPegawai2 & "'")
    If rs.EOF = False Then tempKdJenisPegawai2 = rs("KdJenisPegawai") Else tempKdJenisPegawai2 = ""
    
    Call msubRecFO(rs, "select KdJenisPegawai from DataPegawai where IdPegawai='" & tempIdPegawai3 & "'")
    If rs.EOF = False Then tempKdJenisPegawai3 = rs("KdJenisPegawai") Else tempKdJenisPegawai3 = ""
    
    tempTotalTarifPenjamin = tempTarifKelasPenjaminDB + fTarifCito
    Call msubRecFO(rs, "select KdDetailJenisJasaPelayanan from PasienDaftar where NoPendaftaran='" & fNoPendaftaran & "'")
    If rs.EOF = False Then tempKdDetailJenisJasaPelayanan = rs("KdDetailJenisJasaPelayanan")
    
    Call msubRecFO(rs, "select KdJenisPegawai from DataPegawai where IdPegawai='" & fIdPegawai & "'")
    If rs.EOF = False Then tempKdJenisPegawai = rs("KdJenisPegawai") Else tempKdJenisPegawai = ""
    If tempKdJenisPegawai = "001" Then
        tempIdDokter = tempKdJenisPegawai
    Else
        tempIdDokter = ""
    End If

    Call msubRecFO(rs, "select KdPelayananRS from ConvertPelayananToJasaDokter where KdDetailJenisJasaPelayanan='" & tempKdDetailJenisJasaPelayanan & "' and KdPelayananRS='" & fKdPelayananRS & "'")
    If rs.EOF = True Then
        Call msubRecFO(rsLooping, "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas= '" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'")
    Else
        If tempIdDokter = "" Then
            Call msubRecFO(rsLooping, "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas= '" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('02','04','14')")
        End If
        
        If tempIdPegawai2 = "" And tempIdPegawai3 = "" And tempIdDokter <> "" Then
            Call msubRecFO(rsLooping, "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas= '" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen not in ('04','14')")
        End If
    
        If tempIdPegawai2 <> "" And tempIdPegawai3 = "" And tempIdDokter <> "" Then
            Call msubRecFO(rsLooping, "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas= '" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "' and KdKomponen <> '14'")
        End If
    
        If tempIdPegawai2 <> "" And tempIdPegawai3 <> "" And tempIdDokter <> "" Then
            Call msubRecFO(rsLooping, "select KdKomponen from HargaKomponen where KdPelayananRS='" & fKdPelayananRS & "' and KdKelas= '" & fKdKelas & "' and KdJenisTarif='" & fKdJenisTarif & "'")
        End If
    End If

    While rsLooping.EOF = False
        tempKdKomponen = rsLooping("KdKomponen")
        Call msubRecFO(rs, "SELECT dbo.FB_NewTakeTarifBPTMK('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & tempKdKomponen & "') AS Harga")
        If rs.EOF = False Then tempHarga = rs(0) Else tempHarga = 0
        
        tempJmlPembebasanPerKomp = 0
        If tempTarifKelasPenjaminDB = 0 Then
            tempJmlHutangPerKomp = 0
            tempJmlTanggunganPerKomp = 0
        Else
            tempJmlHutangPerKomp = CDec(tempHarga) / CDec(tempTotalTarifPenjamin) * CDec(tempJmlHutangPenjaminDB)
            tempJmlTanggunganPerKomp = (CDec(tempHarga) / CDec(tempTotalTarifPenjamin)) * CDec(tempJmlTanggunganRSDB)
        End If
        
        Call msubRecFO(rs, "select NoPendaftaran from TempHargaKomponen where NoPendaftaran= '" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & tempKdKomponen & "' and NoStruk is null")
        If rs.EOF = True Then
            If tempKdKomponen <> "04" And tempKdKomponen <> "14" Then
                dbConn.Execute "insert into TempHargaKomponen values('" & fNoPendaftaran & "', '" & fKdRuangan & "', '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & tempKdKomponen & "', '" & fKdJenisTarif & "', " & tempHarga & ", " & fJmlPelayanan & ", null, '" & tempIdPegawai1 & "', " & tempJmlHutangPerKomp & " , " & tempJmlTanggunganPerKomp & ", " & tempJmlPembebasanPerKomp & ",null)"
            End If
            If tempKdKomponen = "04" Then
                dbConn.Execute "insert into TempHargaKomponen values('" & fNoPendaftaran & "', '" & fKdRuangan & "', '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & tempKdKomponen & "', '" & fKdJenisTarif & "', " & tempHarga & ", " & fJmlPelayanan & ", null, '" & tempIdPegawai2 & "', " & tempJmlHutangPerKomp & " , " & tempJmlTanggunganPerKomp & ", " & tempJmlPembebasanPerKomp & ",null)"
            End If
            If tempKdKomponen = "14" Then
                dbConn.Execute "insert into TempHargaKomponen values('" & fNoPendaftaran & "', '" & fKdRuangan & "', '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & tempKdKomponen & "', '" & fKdJenisTarif & "', " & tempHarga & ", " & fJmlPelayanan & ", null, '" & tempIdPegawai3 & "', " & tempJmlHutangPerKomp & " , " & tempJmlTanggunganPerKomp & ", " & tempJmlPembebasanPerKomp & ",null)"
            End If
        Else
            If tempKdKomponen <> "04" And tempKdKomponen <> "14" Then
                dbConn.Execute "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "', KdKelas= '" & fKdKelas & "',Harga=" & tempHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & tempIdPegawai1 & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen= '" & tempKdKomponen & "' and NoStruk is null"
            End If
            If tempKdKomponen = "04" Then
                dbConn.Execute "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "', KdKelas= '" & fKdKelas & "',Harga=" & tempHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & tempIdPegawai2 & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen= '" & tempKdKomponen & "' and NoStruk is null"
            End If
            If tempKdKomponen = "14" Then
                dbConn.Execute "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "', KdKelas= '" & fKdKelas & "',Harga=" & tempHarga & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & tempIdPegawai3 & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen= '" & tempKdKomponen & "' and NoStruk is null"
            End If
        End If
    
    '    --execute AM_DataPelayananTMPasienPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponen,@Harga,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,'A'
    
    '    --if @KdJenisPegawai1='001' and tempKdKomponen not in ('04','14','01')
    '    --begin
    '        --execute AM_DataPelayananTMPasienDokterPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponen,@Harga,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,tempIdPegawai1,'A'
    '    --end
    '    --if @KdJenisPegawai2='001' and tempKdKomponen='04'
    '    --begin
    '        --execute AM_DataPelayananTMPasienDokterPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponen,@Harga,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,tempIdPegawai2,'A'
    '    --end
    '    --if @KdJenisPegawai3='001' and tempKdKomponen='14'
    '    --begin
    '       --execute AM_DataPelayananTMPasienDokterPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponen,@Harga,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,tempIdPegawai3,'A'
    '    --end
        rsLooping.MoveNext
    Wend

'--end

'--Tarif Total
'--begin
    Call msubRecFO(rs, "select KdKomponenTarifTotalTM from MasterDataPendukung")
    If rs.EOF = False Then tempKdKomponenTarifTotal = rs("KdKomponenTarifTotalTM") Else tempKdKomponenTarifTotal = "12"

    Call msubRecFO(rs, "select dbo.FB_NewTakeTarifBPTM ('" & fNoPendaftaran & "', '" & fKdPelayananRS & "', '" & fKdKelas & "', '" & fKdJenisTarif & "', '" & fStatusCito & "', '" & tempIdPegawai1 & "', '" & tempIdPegawai2 & "', '" & tempIdPegawai3 & "', 'T') AS TarifTotal")
    If rs.EOF = False Then tempTarifTotal = rs(0)

    Call msubRecFO(rs, "select NoPendaftaran from TempHargaKomponen where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & tempKdKomponenTarifTotal & "' and NoStruk is null")
    If rs.EOF = True Then
        strSQL = "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "','" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "','" & fKdPelayananRS & "','" & fKdKelas & "', '" & tempKdKomponenTarifTotal & "','" & fKdJenisTarif & "'," & tempTarifTotal & ", " & fJmlPelayanan & ",null,'" & tempIdPegawai1 & "'," & tempJmlHutangPenjaminDB & ", " & tempJmlTanggunganRSDB & "," & tempJmlPembebasanDB & ",null)"
        dbConn.Execute strSQL
    Else
        dbConn.Execute "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & tempTarifTotal & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & tempIdPegawai1 & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & tempKdKomponenTarifTotal & "' and NoStruk is null"
    End If

'--execute AM_DataPelayananTMPasienPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponenTarifTotal,@TarifTotal,tempJmlHutangPenjaminDB,@JmlTanggunganRSDB,@JmlPembebasanDB,@KdKelas,'A'

'--end

'--Tarif Cito
'--begin
    If fStatusCito = "1" Then
        Call msubRecFO(rs, "select KdKomponenTarifCito from MasterDataPendukung")
        If rs.EOF = False Then tempKdKomponenTarifCito = rs("KdKomponenTarifCito") Else tempKdKomponenTarifCito = "07"
        tempJmlPembebasanPerKomp = 0
        If tempTarifKelasPenjaminDB = 0 Then
            tempJmlHutangPerKomp = 0
            tempJmlTanggunganPerKomp = 0
        Else
            tempJmlHutangPerKomp = (CDec(fTarifCito) / CDec(tempTotalTarifPenjamin)) * CDec(tempJmlHutangPenjaminDB)
            tempJmlTanggunganPerKomp = (CDec(fTarifCito) / CDec(tempTotalTarifPenjamin)) * CDec(tempJmlTanggunganRSDB)
        End If
        
        Call msubRecFO(rs, "select NoPendaftaran from TempHargaKomponen where NoPendaftaran=@NoPendaftaran and KdRuangan=@KdRuangan and TglPelayanan=@TglPelayanan and KdPelayananRS=@KdPelayananRS and KdKomponen=tempKdKomponenTarifCito and NoStruk is null")
        If rs.EOF = True Then
            dbConn.Execute "insert into TempHargaKomponen values('" & fNoPendaftaran & "','" & fKdRuangan & "', '" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "', '" & fKdPelayananRS & "','" & fKdKelas & "','" & tempKdKomponenTarifCito & "','" & fKdJenisTarif & "'," & fTarifCito & "," & fJmlPelayanan & ",null,'" & tempIdPegawai1 & "'," & tempJmlHutangPerKomp & "," & tempJmlTanggunganPerKomp & "," & tempJmlPembebasanPerKomp & ",null)"
        Else
            dbConn.Execute "update TempHargaKomponen set KdJenisTarif='" & fKdJenisTarif & "',KdKelas='" & fKdKelas & "',Harga=" & fTarifCito & ",JmlPelayanan=" & fJmlPelayanan & ",IdPegawai='" & tempIdPegawai1 & "' where NoPendaftaran='" & fNoPendaftaran & "' and KdRuangan='" & fKdRuangan & "' and TglPelayanan='" & Format(fTglPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and KdPelayananRS='" & fKdPelayananRS & "' and KdKomponen='" & tempKdKomponenTarifCito & "' and NoStruk is null"
        End If
    End If

'    --execute AM_DataPelayananTMPasienPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponenTarifCito,@TarifCito,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,'A'

'    --if @KdJenisPegawai1='001'
'    --begin
'        --execute AM_DataPelayananTMPasienDokterPH @NoPendaftaran,@KdPelayananRS,@TglPelayanan,@KdRuangan,@KdRuanganAsal,tempKdKomponenTarifCito,@TarifCito,@JmlHutangPerKomp,@JmlTanggunganPerKomp,@JmlPembebasanPerKomp,@KdKelas,tempIdPegawai1,'A'
'    --end
'--end
Exit Function
errLoad:
    Add_TempHargaKomponen = False
    Call msubPesanError("Add_TempHargaKomponen")
End Function

