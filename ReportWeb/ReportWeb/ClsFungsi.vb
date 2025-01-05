
Imports Microsoft.VisualBasic
Imports MySql.Data.MySqlClient
Imports System.Data
Imports System.Security.Cryptography
Imports System.IO
Imports System.Collections.Generic
Imports System.Data.OleDb
Imports ICSharpCode.SharpZipLib.Core
Imports ICSharpCode.SharpZipLib.Zip
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports ReportWeb.ClsWebVer
Imports System.Net.Mail
Imports System.Net
Imports System.Text
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports System.Web
Imports System.Collections.Specialized
Imports Newtonsoft.Json
Imports System.Xml

Public Class ClsFungsi

#Region "Inisialisasi"

    'Private NamaFileTracelog As String = System.Web.HttpContext.Current.Server.MapPath("~/TrcIRX.txt")
    Private NamaFileTracelog As String = System.Web.HttpContext.Current.Server.MapPath("~/Tracelog/TrcIRX_" & Date.Now.ToString("yyyyMMdd") & ".txt")
    Dim ObjCon As New ClsConnection
    Dim ObjSQL As New ClsSQL
    Private SqlParam As New Dictionary(Of String, String)
    Private MasterMCon As MySqlConnection
    Private IsTlsGoogleApi As Boolean = True

#End Region

#Region "Const Datatable"

    Public Function JenisPenerimaDt() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"DCI", "DC Indomaret"})
            dt.Rows.Add(New Object() {"EXP", "Ekspedisi"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function KewajaranDt() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"", "ALL"})
            dt.Rows.Add(New Object() {"WAJAR", "WAJAR"})
            dt.Rows.Add(New Object() {"TIDAK WAJAR", "TIDAK WAJAR"})
            dt.Rows.Add(New Object() {"SANGAT TIDAK WAJAR", "SANGAT TIDAK WAJAR"})
            dt.Rows.Add(New Object() {"BELUM SELESAI", "BELUM SELESAI"})
            dt.Rows.Add(New Object() {"SELESAI", "SELESAI"})
            dt.Rows.Add(New Object() {"GAGAL|HILANG RUSAK BENCANA", "GGL & HRB"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function PerformaDt() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"", "ALL"})
            dt.Rows.Add(New Object() {"TEPAT WAKTU", "TEPAT WAKTU"})
            dt.Rows.Add(New Object() {"TELAT", "TELAT"})
            dt.Rows.Add(New Object() {"RETUR", "RETUR"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function KewajaranDt2() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"", "ALL"})
            dt.Rows.Add(New Object() {"TIDAK WAJAR", "TIDAK WAJAR"})
            dt.Rows.Add(New Object() {"SANGAT TIDAK WAJAR", "SANGAT TIDAK WAJAR"})
            dt.Rows.Add(New Object() {"UNDEFINED", "UNDEFINED"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function ChangeType() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Value", GetType(String))
            dt.Columns.Add("Display", GetType(String))

            dt.Rows.Add(New Object() {"", ""})
            dt.Rows.Add(New Object() {"EMAIL", "EMAIL"})
            dt.Rows.Add(New Object() {"STORECODE", "KODE TOKO"})
            dt.Rows.Add(New Object() {"COPHONE", "NO HP PENERIMA"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    'Public Function PenyebabDt() As DataTable

    '    Dim dt As New DataTable

    '    Try

    '        dt.Columns.Add("Value", GetType(String))
    '        dt.Columns.Add("Display", GetType(String))

    '        dt.Rows.Add(New Object() {"", ""})
    '        dt.Rows.Add(New Object() {"HILANG", "HILANG (HRB)"})
    '        dt.Rows.Add(New Object() {"RUSAK", "RUSAK (HRB)"})
    '        dt.Rows.Add(New Object() {"BENCANA", "BENCANA (HRB)"})
    '        dt.Rows.Add(New Object() {"GAGAL KIRIM", "GAGAL KIRIM (GGL)"})
    '        dt.Rows.Add(New Object() {"KELEBIHAN FISIK", "KELEBIHAN FISIK (GGL)"})
    '        dt.Rows.Add(New Object() {"RETUR DI SERAHKAN TANPA PIN", "RETUR DI SERAHKAN TANPA PIN (HRB)"})
    '        dt.Rows.Add(New Object() {"TIDAK BOLEH AOK", "TIDAK BOLEH AOK"})
    '        'dt.Rows.Add(New Object() {"RESET TIDAK BISA AOK", "RESET TIDAK BISA AOK"}) 'tidak boleh diaktifkan, karena ada ketentuan pendataan dengan Shopee
    '        dt.Rows.Add(New Object() {"DELAWBINSJDC", "HAPUS DARI SURAT JALAN MOBIL DC / ROTI"}) 'paket mau di GGL / HRB, tapi ada di Surat Jalan Mobil DC / Roti

    '        Return dt

    '    Catch ex As Exception

    '        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '        Dim Pesan As String = ""
    '        Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '        WriteTracelogTxt(Pesan)

    '        Return Nothing

    '    Finally

    '        dt = Nothing
    '        GC.Collect()

    '    End Try

    'End Function

    'Public Function PenyebabDtV2() As DataTable

    '    Dim dt As New DataTable

    '    Try

    '        dt.Columns.Add("Value", GetType(String))
    '        dt.Columns.Add("Display", GetType(String))
    '        dt.Columns.Add("Status", GetType(String))

    '        dt.Rows.Add(New Object() {"", "", ""})

    '        dt.Rows.Add(New Object() {"TOKO SERAH TANPA PIN", "TOKO SERAH TANPA PIN (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"TOKO SALAH BUAT AWB", "TOKO SALAH BUAT AWB (GGL)", "FLR"})
    '        dt.Rows.Add(New Object() {"TOKO SALAH TIMBANG UKUR", "TOKO SALAH TIMBANG UKUR (GGL)", "FLR"})
    '        dt.Rows.Add(New Object() {"TOKO RETUR TANPA FISIK", "TOKO RETUR TANPA FISIK (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"HILANG DI TOKO", "HILANG DI TOKO (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"RUSAK DI TOKO", "RUSAK DI TOKO (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"HILANG DI DRIVER", "HILANG DI DRIVER (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"RUSAK DI DRIVER", "RUSAK DI DRIVER (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"HILANG DI DC", "HILANG DI DC (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"RUSAK DI DC", "RUSAK DI DC (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"ADMIN HO GAGAL BUAT AWB", "ADMIN HO GAGAL BUAT AWB (GGL)", "FLR"})
    '        dt.Rows.Add(New Object() {"BENCANA", "BENCANA (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"TRIAL", "TRIAL (GGL)", "FLR"})
    '        dt.Rows.Add(New Object() {"TIDAK BOLEH AOK", "TIDAK BOLEH AOK", "XDO"})
    '        'dt.Rows.Add(New Object() {"RESET TIDAK BISA AOK", "RESET TIDAK BISA AOK", "YDO"})'tidak boleh diaktifkan, karena ada ketentuan pendataan dengan Shopee
    '        dt.Rows.Add(New Object() {"HILANG DI 3PL", "HILANG DI 3PL (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"RUSAK DI 3PL", "RUSAK DI 3PL (HRB)", "OOC"})
    '        dt.Rows.Add(New Object() {"PARTNER SALAH BUAT AWB", "PARTNER SALAH BUAT AWB (GGL)", "FLR"})
    '        dt.Rows.Add(New Object() {"DELAWBINSJDC", "HAPUS DARI SURAT JALAN MOBIL DC / ROTI", "DELAWBINSJDC"}) 'paket mau di GGL / HRB, tapi ada di Surat Jalan Mobil DC / Roti

    '        Return dt

    '    Catch ex As Exception

    '        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '        Dim Pesan As String = ""
    '        Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '        WriteTracelogTxt(Pesan)

    '        Return Nothing

    '    Finally

    '        dt = Nothing
    '        GC.Collect()

    '    End Try

    'End Function

    Public Function PenyebabDtV2() As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `report_flaghrb_reasonlist`();"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

#End Region

#Region "Tracelog"

    Public Sub WriteTracelog_(ByVal Pesan As String)

        Dim PesanError As String = ""

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Try

            SqlQuery = "INSERT INTO xxxTracelogxxx ("
            SqlQuery &= vbCrLf & " AddTime, AddID, Message"
            SqlQuery &= vbCrLf & " ) VALUES ("
            SqlQuery &= vbCrLf & " NOW(), USER(), @Pesan"
            SqlQuery &= vbCrLf & " )"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Pesan", Pesan)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam, PesanError)

        Catch ex As Exception

            Dim str As String = "TraceLog:" & Pesan
            str &= vbCrLf & "Catch Error:" & ex.Message

            WriteTracelogTxt("Error :" & ex.Message & vbCrLf & "Tracelog :" & Pesan)

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

    End Sub

    Public Sub WriteTracelogTxt(ByVal Pesan As String)

        'Dim stackTrace As Diagnostics.StackTrace = New Diagnostics.StackTrace()
        'Dim NamaFungsi = stackTrace.GetFrame(1).GetMethod().Name
        'Dim str As String = "Tanggal Tracelog:" & Date.Now & vbCrLf & "Fungsi:" & NamaFungsi.ToString & vbCrLf & Pesan & vbCrLf & vbCrLf

        'Dim objWriter As New System.IO.StreamWriter(NamaFileTracelog, True)
        'objWriter.WriteLine(str & vbCrLf)
        'objWriter.Close()

        Dim str As String = "Tanggal Tracelog:" & Date.Now & vbCrLf & Pesan & vbCrLf & vbCrLf

        Dim MapPath As String = System.Web.HttpContext.Current.Server.MapPath("~/Tracelog/")

        If Not Directory.Exists(MapPath) Then
            Directory.CreateDirectory(MapPath)
        End If

        Using objWriter As New System.IO.StreamWriter(NamaFileTracelog, True)
            objWriter.Write(str)
        End Using

    End Sub

#End Region

#Region "Cek String"

    Public Function isNumberOnly(ByVal input As String) As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^[0-9]+$")

    End Function

    Public Function isAlphaOnly(ByVal input As String) As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^[a-zA-Z ]+$")

    End Function

    Public Function isAlphaNumeric(ByVal input As String) As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^[a-zA-Z 0-9]+$")

    End Function

    Public Function isAlphaAndNumeric(ByVal input As String) As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$")

    End Function

    Public Function isDouble(ByVal input As String) As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^[0-9]+(\.[0-9]+)?$")

    End Function

    Public Function isRegexTrue(ByVal input As String, Optional ByVal regex As String = "") As Boolean

        Return System.Text.RegularExpressions.Regex.IsMatch(input, "^[a-zA-Z 0-9" & regex & "]+$")

    End Function

#End Region

#Region "Response Redirect"

    Public Function RequestFromString(ByVal key As String, Optional ByVal CaseSensitive As Boolean = False) As String

        Try

            Dim page As String = HttpContext.Current.Request.Url.AbsoluteUri

            If page.Contains("?") Then
                Dim curPage As String = page.Split("?")(0)
                Dim reqQuery As String = page.Split("?")(1)
                reqQuery = Decrypt(reqQuery, "x")

                Dim hasil As String = ""

                If (reqQuery & "") <> "" Then

                    For Each i As String In reqQuery.Split("&")
                        Dim param As String = i.Split("=")(0)
                        Dim value As String = i.Split("=")(1)
                        If CaseSensitive Then

                            If param = key Then 'RequestQuery String tipe yg sensitif --> huruf gede kecil pengaruh
                                hasil = value
                                Return hasil
                            End If

                        Else

                            If param.ToLower = key.ToLower Then
                                hasil = value
                                Return hasil
                            End If

                        End If
                    Next

                End If

                Return hasil

            Else

                Return ""

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Function _RedirectEncrypt(ByVal url As String) As String

        Try

            If url.Contains("?") Then
                Dim CurPage As String = url.Split("?")(0)
                Dim DirectedPageValue As String = url.Split("?")(1)
                DirectedPageValue = Encrypt(DirectedPageValue, "x")
                Return CurPage & "?" & DirectedPageValue
            Else
                Return url
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Dim page As String = HttpContext.Current.Request.Url.AbsoluteUri
            Return page

        End Try

    End Function

#End Region

#Region "Encrypt Decrypt"

    Private Function MD5Hash(ByVal value As String) As Byte()
        Dim MD5 As New MD5CryptoServiceProvider
        Return MD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(value))
    End Function

    Public Function Encrypt(ByVal stringToEncrypt As String, ByVal key As String) As String

        Try

            Dim DES As New TripleDESCryptoServiceProvider
            Dim MD5 As New MD5CryptoServiceProvider

            DES.Key = MD5Hash(key)
            DES.Mode = CipherMode.ECB
            Dim Buffer As Byte() = ASCIIEncoding.ASCII.GetBytes(stringToEncrypt)
            Return Convert.ToBase64String(DES.CreateEncryptor().TransformFinalBlock(Buffer, 0, Buffer.Length))

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

    Public Function Decrypt(ByVal encryptedString As String, ByVal key As String) As String

        Try

            Dim DES As New TripleDESCryptoServiceProvider
            Dim MD5 As New MD5CryptoServiceProvider

            DES.Key = MD5Hash(key)
            DES.Mode = CipherMode.ECB
            Dim Buffer As Byte() = Convert.FromBase64String(encryptedString)
            Return ASCIIEncoding.ASCII.GetString(DES.CreateDecryptor().TransformFinalBlock(Buffer, 0, Buffer.Length))

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

#End Region

#Region "Get Data"

    Public Function GetProv(ByVal MCon As MySqlConnection) As DataTable

        Dim dt As DataTable

        Dim SqlQuery As String

        Try

            SqlQuery = "SELECT Province AS Kode, CONCAT(Name, '(', Alias, ')') AS Nama"
            SqlQuery &= " FROM MstProvince"

            Dim SqlParam As New Dictionary(Of String, String)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetKota(ByVal MCon As MySqlConnection, Optional ByVal Provinsi As String = "") As DataTable

        Dim SqlQuery As String

        Try

            SqlQuery = "SELECT City AS Kode, CONCAT(Name, '(', Alias, ')') AS Nama"
            SqlQuery &= " FROM MstCity"
            If Provinsi <> "" Then
                SqlQuery &= vbCrLf & " WHERE Province = @Provinsi"
            End If

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@Provinsi", Provinsi)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            'Dim dt As DataTable = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            'Return dt

            Return ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

    Public Function GetStation(ByVal MCon As MySqlConnection) As DataTable

        Dim SqlQuery As String

        Try

            SqlQuery = "SELECT Station AS Kode, CONCAT(Name, '(', Alias, ')') AS Nama"
            SqlQuery &= " FROM MstStation"
            SqlQuery &= " WHERE ActiveDate <= CURDATE() AND (InactiveDate >= CURDATE() OR InactiveDate IS NULL)"

            Dim SqlParam As New Dictionary(Of String, String)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            'Dim dt As DataTable = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            'Return dt

            Return ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

    Public Function GetJenisRute() As DataTable

        Dim dt As New DataTable

        Try

            dt.Columns.Add("Kode", GetType(String))
            dt.Columns.Add("Nama", GetType(String))

            dt.Rows.Add(New Object() {"STA", "Antar Station"})
            dt.Rows.Add(New Object() {"HUB", "Station - Hub"})

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetRack(ByVal MCon As MySqlConnection, ByVal HubCode As String) As DataTable

        Dim SqlQuery As String

        Try

            SqlQuery = "SELECT Rack AS Kode, Name AS Nama"
            SqlQuery &= " FROM MstRack"
            SqlQuery &= " WHERE Hub = @HubCode"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@HubCode", HubCode)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            'Dim dt As DataTable = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            'Return dt

            Return ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

    Public Function GetMyHub(ByVal MCon As MySqlConnection, ByVal HubCode As String)

        Dim SqlQuery As String

        Try

            SqlQuery = "SELECT *"
            SqlQuery &= " FROM MstHub"
            SqlQuery &= " WHERE Hub = @HubCode"

            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@HubCode", HubCode)

            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            'Dim dt As DataTable = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            'Return dt

            Return ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

#End Region

#Region "LAIN2"

    Public Function ReadWebConfig(ByVal NamaSetting As String, Optional ByVal isMapPath As Boolean = False) As String

        Try

            Dim ValueSetting As String = "" & ConfigurationManager.AppSettings(NamaSetting)

            If isMapPath Then
                ValueSetting = System.Web.HttpContext.Current.Server.MapPath(ValueSetting)
            End If

            Return ValueSetting

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Function CekApakahInputDenganKetik(ByVal Awal As DateTime, ByVal Akhir As DateTime, Optional ByVal IntervalMilliseconds As Integer = 0) As Boolean
        Dim BatasWaktuMax As Integer = 500 'dalam Milliseconds

        If IntervalMilliseconds > 0 Then
            BatasWaktuMax = IntervalMilliseconds
        End If

        Dim ts As TimeSpan
        ts = Akhir.Subtract(Awal)

        If ts.TotalMilliseconds <= BatasWaktuMax Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function _CleanParamSql(ByVal input As String) As String

        Dim hasil As String = ""
        Try
            hasil = input

            hasil = hasil.Replace("'", "''")
            hasil = hasil.Replace("\", "\\")
            'hasil = hasil.Replace(""", """")

            Return hasil

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return input

        End Try

    End Function

    Public Sub CreateMsgBox(ByVal Smsg As String, Optional ByVal url As String = "")

        Dim curPage As Page = HttpContext.Current.CurrentHandler

        If Not IsNothing(curPage) Then
            Dim script As ClientScriptManager = curPage.ClientScript

            If Not script.IsStartupScriptRegistered(Me.GetType, "Alert") Then
                Dim urlredirect As String = ""
                If url <> "" Then
                    urlredirect = "window.location.href='" & url & "';"
                End If
                script.RegisterStartupScript(Me.GetType, "Alert", "<script type=text/javascript>alert('" & Smsg & "');" & urlredirect & "</script>")
            End If

        End If

    End Sub


    Public Function ConvertStringToDatatable(ByVal Param As String) As DataTable

        Dim SplitterRow(0) As String
        SplitterRow(0) = "|"

        Dim SplitterColumn(0) As String
        SplitterColumn(0) = ","

        Return ConvertStringToDatatable(Param, SplitterRow, SplitterColumn)

    End Function

    Public Function ConvertStringToDatatable(ByVal Param As String, ByVal SplitterRow As String(), ByVal SplitterColumn As String()) As DataTable

        Dim dtResult As DataTable = Nothing

        Try

            Dim Result As String() = Param.ToString.Split(SplitterRow, StringSplitOptions.None) 'pemisah row
            For i As Integer = 0 To Result.Length - 1
                If i = 0 Then 'column header
                    If dtResult Is Nothing Then
                        dtResult = New DataTable
                    End If
                Else 'data
                    dtResult.Rows.Add()
                End If
                Dim Result2Split As String() = Result(i).Split(SplitterColumn, StringSplitOptions.None) 'pemisah kolom
                For j As Integer = 0 To Result2Split.Length - 1
                    If i = 0 Then 'column header
                        dtResult.Columns.Add(Result2Split(j))
                    Else 'data
                        dtResult.Rows(i - 1).Item(j) = Result2Split(j)
                    End If
                Next
            Next

            Return dtResult

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dtResult

        Finally

            dtResult = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function ConvertStringToDatatableWithBarcode(ByVal Param As String) As DataTable

        Dim dtResult As DataTable = Nothing

        Try

            Dim BarisSplit As String() = New String() {"#|#"}
            Dim Result As String() = Param.ToString.Split(BarisSplit, StringSplitOptions.None) 'pemisah row
            For i As Integer = 0 To Result.Length - 1
                If i = 0 Then 'column header
                    If dtResult Is Nothing Then
                        dtResult = New DataTable
                    End If
                Else 'data
                    dtResult.Rows.Add()
                End If
                Dim KolomSplit As String() = New String() {"#,#"}
                Dim Result2Split As String() = Result(i).Split(KolomSplit, StringSplitOptions.None) 'pemisah kolom
                For j As Integer = 0 To Result2Split.Length - 1
                    If i = 0 Then 'column header
                        dtResult.Columns.Add(Result2Split(j))
                    Else 'data
                        dtResult.Rows(i - 1).Item(j) = Result2Split(j)
                    End If
                Next
            Next

            Return dtResult

        Catch ex As Exception

            Return dtResult

        Finally

            dtResult = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function ConvertDatatableToString(ByVal Param As DataTable, ByVal SplitColumn As String, ByVal SplitLine As String) As String

        Dim Result As String = ""

        Try

            For i As Integer = 0 To Param.Rows.Count - 1

                If i > 0 Then
                    Result &= SplitLine
                End If

                For j As Integer = 0 To Param.Columns.Count - 1
                    If j > 0 Then
                        Result &= SplitColumn
                    End If
                    Result &= Param.Rows(i).Item(j).ToString
                Next

            Next

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Result = Nothing

        End Try

        Return Result

    End Function


    Public Function DataTableToCSV(ByVal datatable As DataTable, ByVal separator As Char, ByVal UseHeader As Boolean) As String

        Return DataTableToCSV(datatable, separator, "", UseHeader)

    End Function

    Public Function DataTableToCSV(ByVal datatable As DataTable, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean) As String

        Dim sb As New StringBuilder

        Try

            If UseHeader Then
                For i As Integer = 0 To datatable.Columns.Count - 1
                    If i > 0 Then
                        sb.Append(separator)
                    End If
                    sb.Append(text_delimiter & datatable.Columns(i).ColumnName & text_delimiter)
                Next
                sb.AppendLine()
            End If

            For Each dr As DataRow In datatable.Rows
                For i As Integer = 0 To datatable.Columns.Count - 1
                    If i > 0 Then
                        sb.Append(separator)
                    End If
                    sb.Append(text_delimiter & dr(i) & text_delimiter)
                Next
                sb.AppendLine()
            Next

            'Dim Hasil As String = sb.ToString
            Return sb.ToString

        Catch ex As Exception

            Return sb.ToString

        Finally

            sb = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function DataSetToCSV(ByVal dataset As DataSet, ByVal separator As Char, ByVal UseHeader As Boolean) As String

        Dim sb As New StringBuilder

        Try

            For Each datatable As DataTable In dataset.Tables

                If UseHeader Then
                    For i As Integer = 0 To datatable.Columns.Count - 1
                        If i > 0 Then
                            sb.Append(separator)
                        End If
                        sb.Append(datatable.Columns(i).ColumnName)
                    Next
                    sb.AppendLine()
                End If

                For Each dr As DataRow In datatable.Rows
                    For i As Integer = 0 To datatable.Columns.Count - 1
                        If i > 0 Then
                            sb.Append(separator)
                        End If
                        sb.Append(("" & dr(i)))
                    Next
                    sb.AppendLine()
                Next

            Next

            Return sb.ToString

        Catch ex As Exception

            Return sb.ToString

        Finally

            sb = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function DataTableToFile(ByVal datatable As DataTable, ByVal separator As Char, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataTableToFile(datatable, separator, "", UseHeader, FileName, ExtFile, FolderName, False, False, "", PesanError)

    End Function

    Public Function DataTableToFile(ByVal datatable As DataTable, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataTableToFile(datatable, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, False, False, "", PesanError)

    End Function

    Public Function DataTableToFile(ByVal datatable As DataTable, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByRef PesanError As String) As String

        Return DataTableToFile(datatable, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, False, "", PesanError)

    End Function

    Public Function DataTableToFile(ByVal datatable As DataTable, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByVal AppendFile As Boolean, ByRef PesanError As String) As String

        Return DataTableToFile(datatable, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, AppendFile, "", PesanError)

    End Function

    Public Function DataTableToFile(ByVal datatable As DataTable, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByVal AppendFile As Boolean, ByVal CopyFolderPath As String, ByRef PesanError As String) As String

        Dim DsTemp As New DataSet

        DsTemp.Tables.Add(datatable)

        Return DataSetToFile(DsTemp, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, AppendFile, CopyFolderPath, PesanError)

    End Function

    Public Function DataSetToFile(ByVal dataset As DataSet, ByVal separator As Char, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataSetToFile(dataset, separator, "", UseHeader, FileName, ExtFile, FolderName, False, True, "", PesanError)

    End Function

    Public Function DataSetToFile(ByVal dataset As DataSet, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataSetToFile(dataset, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, False, True, "", PesanError)

    End Function

    Public Function DataSetToFile(ByVal dataset As DataSet, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByRef PesanError As String) As String

        Return DataSetToFile(dataset, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, True, "", PesanError)

    End Function

    Public Function DataSetToFile(ByVal dataset As DataSet, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByVal AppendFile As Boolean, ByRef PesanError As String) As String

        Return DataSetToFile(dataset, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, AppendFile, "", PesanError)

    End Function

    Public Function DataSetToFile(ByVal dataset As DataSet, ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByVal AppendFile As Boolean, ByVal CopyFolderPath As String, ByRef PesanError As String) As String

        Dim Result As String = "0|Failed"

        Try

            If IsNothing(dataset) Then

                Result = "0|DataSet Is Nothing!"

                PesanError = "DataSet Is Nothing!"
                Return Result

            Else

                If dataset.Tables.Count < 1 Then

                    Result = "0|No DataTable in DataSet!"

                    PesanError = "No DataTable in DataSet!"
                    Return Result

                End If

                Dim dtIsNothing As Boolean = True
                Dim dtIsEmpty As Boolean = True

                For Each datatable As DataTable In dataset.Tables

                    If Not IsNothing(datatable) Then
                        dtIsNothing = False
                        If datatable.Rows.Count > 0 Then
                            dtIsEmpty = False
                        End If
                    End If

                Next

                If dtIsNothing Then

                    Result = "0|DataTable Is Nothing!"

                    PesanError = "DataTable Is Nothing!"
                    Return Result

                End If

                If dtIsEmpty Then

                    If Not UseHeader Then

                        Result = "0|Tidak ada data!"

                        PesanError = "Tidak ada data!"
                        Return Result

                    End If

                End If

            End If

            Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\GeneratedReport\" & FolderName & "\")
            If FolderName.Trim = "" Then
                Path = Path.Replace("GeneratedReport\\", "GeneratedReport\")
            End If

            Dim UserDirectory As Boolean = False
            Try
                If Directory.Exists(Path) = True Then
                    UserDirectory = True
                Else
                    Directory.CreateDirectory(Path)
                    UserDirectory = True
                End If
            Catch ex As Exception
                WriteTracelogTxt(ex.Message)
                PesanError = ex.Message
            End Try

            If UserDirectory Then

                ExtFile = ExtFile.Trim
                If ExtFile <> "" Then
                    While ExtFile.StartsWith(".")
                        ExtFile = ExtFile.Substring(1, ExtFile.Length - 1)
                    End While
                    ExtFile = "." & ExtFile
                End If
                ExtFile = ExtFile.ToLower

                Dim FullPath As String = Path & FileName & ExtFile
                FullPath = FullPath.Replace("\\" & FileName & ExtFile, "\" & FileName & ExtFile)

                Dim FileIsCreated As Boolean = False

                For Each datatable As DataTable In dataset.Tables

                    Using sw As New StreamWriter(FullPath, AppendFile)

                        If UseHeader Then

                            Dim SBHeader As New StringBuilder
                            Dim HeaderName As String = ""

                            For i As Integer = 0 To datatable.Columns.Count - 1

                                If i > 0 Then
                                    SBHeader.Append(separator)
                                End If

                                HeaderName = datatable.Columns(i).ColumnName

                                If RemoveBold Then
                                    HeaderName = HeaderName.Replace("<b>", "").Replace("</b>", "")
                                End If

                                SBHeader.Append(text_delimiter & HeaderName & text_delimiter)

                            Next

                            sw.WriteLine(SBHeader.ToString)
                            FileIsCreated = True
                            SBHeader = Nothing

                        End If

                        Dim sb As New StringBuilder
                        Dim CellValue As String = ""

                        For Each dr As DataRow In datatable.Rows

                            sb = New StringBuilder

                            For i As Integer = 0 To datatable.Columns.Count - 1

                                If i > 0 Then
                                    sb.Append(separator)
                                End If

                                CellValue = ("" & dr(i)).ToString

                                If RemoveBold Then
                                    CellValue = CellValue.Replace("<b>", "").Replace("</b>", "")
                                End If

                                sb.Append(text_delimiter & CellValue & text_delimiter)

                            Next

                            sw.WriteLine(sb.ToString)
                            FileIsCreated = True
                            sb = Nothing

                        Next

                    End Using

                Next

                If FileIsCreated Then
                    Result = "1|" & FullPath

                    Try
                        If CopyFolderPath <> "" Then

                            CopyFolderPath.Replace("\\", "\")

                            If Not CopyFolderPath.EndsWith("\") Then
                                CopyFolderPath &= "\"
                            End If

                            If Directory.Exists(CopyFolderPath) Then
                                'FileSystem.FileCopy(FullPath, CopyFolderPath & FileName & ExtFile)
                                Dim file = New FileInfo(FullPath)
                                file.CopyTo(CopyFolderPath & FileName & ExtFile, True)
                            Else
                                WriteTracelogTxt("Path Folder Copy tidak ditemukan " & CopyFolderPath)
                                PesanError = "Path Folder Copy tidak ditemukan " & CopyFolderPath
                            End If

                        End If 'dari If CopyFolderPath <> ""

                    Catch ex As Exception
                        WriteTracelogTxt("Gagal CopyFolderPath " & CopyFolderPath & " " & ex.Message)
                        PesanError = "Gagal CopyFolderPath " & CopyFolderPath & " " & ex.Message
                    End Try

                End If 'dari If FileIsCreated

            Else

                PesanError = "User Directory " & Path & " Not Exists"
                Return ""

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            PesanError = ex.Message

        Finally

            GC.Collect()

        End Try

        Return Result

    End Function


    Public Function DataReaderToFile(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal SQLParameters As Dictionary(Of String, String), ByVal separator As Char, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataReaderToFile(Mcon, SQLQuery, SQLParameters, separator, "", UseHeader, FileName, ExtFile, FolderName, False, False, PesanError)

    End Function

    Public Function DataReaderToFile(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal SQLParameters As Dictionary(Of String, String), ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByRef PesanError As String) As String

        Return DataReaderToFile(Mcon, SQLQuery, SQLParameters, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, False, False, PesanError)

    End Function

    Public Function DataReaderToFile(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal SQLParameters As Dictionary(Of String, String), ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByRef PesanError As String) As String

        Return DataReaderToFile(Mcon, SQLQuery, SQLParameters, separator, text_delimiter, UseHeader, FileName, ExtFile, FolderName, RemoveBold, False, PesanError)

    End Function

    Public Function DataReaderToFile(ByVal Mcon As MySqlConnection, ByVal SQLQuery As String, ByVal SQLParameters As Dictionary(Of String, String), ByVal separator As Char, ByVal text_delimiter As String, ByVal UseHeader As Boolean, ByVal FileName As String, ByVal ExtFile As String, ByVal FolderName As String, ByVal RemoveBold As Boolean, ByVal AppendFile As Boolean, ByRef PesanError As String) As String

        Dim Result As String = "0|Failed"

        'Dim Mcom As MySqlCommand = Nothing

        Try

            Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\GeneratedReport\" & FolderName & "\")
            If FolderName.Trim = "" Then
                Path = Path.Replace("GeneratedReport\\", "GeneratedReport\")
            End If

            Dim UserDirectory As Boolean = False
            Try
                If Directory.Exists(Path) = True Then
                    UserDirectory = True
                Else
                    Directory.CreateDirectory(Path)
                    UserDirectory = True
                End If
            Catch ex As Exception
                WriteTracelogTxt(ex.Message)
                PesanError = ex.Message
            End Try

            If UserDirectory Then

                ExtFile = ExtFile.Trim
                If ExtFile <> "" Then
                    While ExtFile.StartsWith(".")
                        ExtFile = ExtFile.Substring(1, ExtFile.Length - 1)
                    End While
                    ExtFile = "." & ExtFile
                End If
                ExtFile = ExtFile.ToLower

                Dim FullPath As String = Path & FileName & ExtFile
                FullPath = FullPath.Replace("\\" & FileName & ExtFile, "\" & FileName & ExtFile)

                Dim FileIsCreated As Boolean = False

                'Pengganti DataSet karena OutOfMemory, langsung dari DataReader
                Using Mcom As MySqlCommand = New MySqlCommand("", Mcon)

                    Mcom.CommandTimeout = 0

                    If Mcon.State <> ConnectionState.Open Then
                        Mcon.Open()
                    End If

                    Mcom.CommandText = SQLQuery

                    If Not IsNothing(SQLParameters) Then

                        For Each Parameter As KeyValuePair(Of String, String) In SQLParameters

                            Mcom.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                        Next

                    End If

                    Using MySqlReader As MySqlDataReader = Mcom.ExecuteReader()

                        Using sw As New StreamWriter(FullPath, AppendFile)

                            Do

                                'Header Table
                                If UseHeader Then

                                    Dim SBHeader As New StringBuilder
                                    Dim HeaderName As String = ""

                                    For i As Integer = 0 To MySqlReader.FieldCount - 1

                                        If i > 0 Then
                                            SBHeader.Append(separator)
                                        End If

                                        If Not IsNothing(MySqlReader.GetName(i)) Then
                                            HeaderName = MySqlReader.GetName(i)
                                        End If

                                        If RemoveBold Then
                                            HeaderName = HeaderName.Replace("<b>", "").Replace("</b>", "")
                                        End If

                                        SBHeader.Append(text_delimiter & HeaderName & text_delimiter)

                                    Next

                                    sw.WriteLine(SBHeader.ToString)
                                    FileIsCreated = True
                                    SBHeader = Nothing

                                End If

                                'Isi Table
                                Dim sb As New StringBuilder
                                Dim CellValue As String = ""

                                While MySqlReader.Read

                                    sb = New StringBuilder

                                    For i As Integer = 0 To MySqlReader.FieldCount - 1

                                        If i > 0 Then
                                            sb.Append(separator)
                                        End If

                                        CellValue = ("" & MySqlReader.GetValue(i)).ToString

                                        If RemoveBold Then
                                            CellValue = CellValue.Replace("<b>", "").Replace("</b>", "")
                                        End If

                                        sb.Append(text_delimiter & CellValue & text_delimiter)

                                    Next

                                    sw.WriteLine(sb.ToString)
                                    FileIsCreated = True
                                    sb = Nothing

                                End While

                            Loop While MySqlReader.NextResult

                        End Using

                    End Using

                End Using

                If FileIsCreated Then
                    Result = "1|" & FullPath
                End If

            Else

                PesanError = "User Directory " & Path & " Not Exists"
                Return ""

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            PesanError = ex.Message

        Finally

            GC.Collect()

        End Try

        Return Result

    End Function


    Public Function ConvertArray2DToString(ByVal MyArray(,) As String) As String

        Dim sb As New StringBuilder

        Try

            For Row As Integer = MyArray.GetLowerBound(0) To MyArray.GetUpperBound(0)

                For Column As Integer = MyArray.GetLowerBound(1) To MyArray.GetUpperBound(1)
                    If Column <> MyArray.GetLowerBound(1) Then
                        sb.Append(",")
                    End If
                    sb.Append(MyArray(Row, Column))
                Next

                If Row <> MyArray.GetUpperBound(0) Then
                    sb.Append("|")
                End If

            Next

            Return sb.ToString

        Catch ex As Exception

            Return sb.ToString

        Finally

            sb = Nothing
            GC.Collect()

        End Try

    End Function

    Public Sub ddlBound(ByVal ddl As DropDownList, ByVal dt As DataTable, ByVal DataText As String, ByVal DataValue As String)

        Try

            ddl.DataSource = dt
            ddl.DataTextField = dt.Columns(DataText).ToString
            ddl.DataValueField = dt.Columns(DataValue).ToString
            ddl.DataBind()

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

    End Sub

    Public Sub rblBound(ByVal rbl As RadioButtonList, ByVal dt As DataTable, ByVal DataText As String, ByVal DataValue As String)

        Try

            rbl.DataSource = dt
            rbl.DataTextField = dt.Columns(DataText).ToString
            rbl.DataValueField = dt.Columns(DataValue).ToString
            rbl.DataBind()

            Try
                rbl.SelectedIndex = 0
            Catch ex As Exception

            End Try

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

    End Sub

    Public Sub chkBound(ByVal chk As CheckBoxList, ByVal dt As DataTable, ByVal DataText As String, ByVal DataValue As String)

        Try

            chk.DataSource = dt
            chk.DataTextField = dt.Columns(DataText).ToString
            chk.DataValueField = dt.Columns(DataValue).ToString
            chk.DataBind()

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

    End Sub

    Public Function HapusDariDataTable(ByVal dt As DataTable, ByVal IndexToBeDelete As Integer) As DataTable

        Dim dtTemp As DataTable

        Try

            dtTemp = dt

            dtTemp.Rows.Remove(dtTemp.Rows(IndexToBeDelete))

            Return dtTemp

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dtTemp = Nothing
            GC.Collect()

        End Try

    End Function

    Private Function CreateFunctionColumn(ByVal dt As DataTable) As String

        Try

            Dim _dtList As DataTable = dt
            Dim TemplateRowFilter As String = ""

            'format: string.format("At {0} in {1}, the temperature was {2} degrees.", dat, city, temp)
            For i As Integer = 0 To dt.Columns.Count - 1
                If _dtList.Columns(i).DataType Is Type.GetType("System.String") Then
                    If TemplateRowFilter <> "" Then
                        TemplateRowFilter &= " or "
                    End If
                    TemplateRowFilter &= "[" & _dtList.Columns(i).ColumnName & "] like '%{0}%'"
                End If
            Next
            If TemplateRowFilter.Trim <> "" Then
                TemplateRowFilter = " 1=1 and (" & TemplateRowFilter & ")"
            End If

            Return TemplateRowFilter

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Function ConditionToString(ByVal ConditionCode As String) As String

        'IIf(Eval("Condition").ToString() = "0", "OK", IIf(Eval("Condition").ToString() = "1", "Sobek", IIf(Eval("Condition").ToString() = "2", "Basah", "")))

        Try
            If ConditionCode = "0" Then

                Return "OK"

            ElseIf ConditionCode = "1" Then

                Return "SOBEK"

            Else

                Return "BASAH"

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Function FileExceltoDataSet(ByVal FilePath As String, ByVal UseSheetsOrder As Boolean, ByVal SheetsOrder() As String, ByRef ErrorMessage As String) As DataSet

        Dim hasil As New DataSet

        Dim MConExcel As New OleDbConnection

        Try

            'MConExcel = ObjCon.SetConnExcelHDR(FilePath) 'Jika Pakai HDR baris pertama jadi Nama Kolom, tapi tipe data baca dari 8 baris pertama, kalau kebetulan angka semua maka baris ke 9 dan seterusnya jadi kosong
            MConExcel = ObjCon.SetConnExcel(FilePath) 'Jika Non HDR baris pertama tetap ada tapi Nama Kolom jadi F1 F2 F3 dst, Efeknya di datatable jadi tidak bisa diambil Nama Kolomnya
            MConExcel.Open()

            'Dim SB As New StringBuilder

            Using dtSheets As DataTable = MConExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)

                'Dim dtSheets As DataTable = New DataTable
                'dtSheets = MConExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)

                'Dim dt As DataTable
                'Dim cmd As OleDbCommand
                'Dim reader As OleDbDataReader

                If UseSheetsOrder Then

                    Dim ExcelTableName As String = ""
                    'Dim DicColumnName As New Dictionary(Of Integer, String)
                    Dim NamaKolom As String = ""

                    For Each TableName As String In SheetsOrder

                        Using dt As DataTable = New DataTable

                            Try

                                ExcelTableName = TableName & "$"

                                Using cmd As OleDbCommand = MConExcel.CreateCommand()

                                    'cmd = MConExcel.CreateCommand()
                                    cmd.CommandText = "SELECT * FROM `" & ExcelTableName & "`;"

                                    Using reader As OleDbDataReader = cmd.ExecuteReader()

                                        'reader = cmd.ExecuteReader()
                                        dt.Load(reader)
                                        dt.TableName = TableName

                                        If dt.Rows.Count > 0 Then

                                            '2022-07-14 15:09, By Cucun, Karena tanggal 13 diubah jadi Non HDR, nama kolom nya jadi F1 F2 F3 dst semua
                                            'Ambil baris pertama sebelum di-delete
                                            'DicColumnName = New Dictionary(Of Integer, String)
                                            For i As Integer = 0 To dt.Columns.Count - 1

                                                'DicColumnName.Add(i, dt.Columns(i).ColumnName)

                                                NamaKolom = "" & dt.Rows(0).Item(i)

                                                If NamaKolom <> "" Then
                                                    dt.Columns(i).ColumnName = NamaKolom
                                                End If

                                            Next

                                            '2022-07-13 18:02, By Cucun, Karena tidak pakai connection string HDR=Y
                                            'Maka harus buang baris pertama manual
                                            'Kalau pakai HDR=Y, ada kasus 8 baris pertama angka, kemudian sisa nya bukan angka, maka jadi kosong
                                            'Dikasih IMEX=1 percuma
                                            dt.Rows(0).Delete()
                                            'Harus pakai AcceptChanges, jika tidak akan error Deleted row information cannot be accessed through the row
                                            dt.AcceptChanges()

                                        End If

                                        hasil.Tables.Add(dt)

                                    End Using

                                End Using

                            Catch ex As Exception
                            End Try

                        End Using

                    Next

                    'DicColumnName = Nothing

                Else

                    Dim TableName As String
                    'Dim DicColumnName As New Dictionary(Of Integer, String)
                    Dim NamaKolom As String = ""

                    For Each dr As DataRow In dtSheets.Rows

                        Using dt As DataTable = New DataTable

                            TableName = "" & dr("TABLE_NAME")

                            If TableName.EndsWith("$") Or TableName.EndsWith("$'") Then

                                TableName = TableName.Replace("'", "")

                                Using cmd As OleDbCommand = MConExcel.CreateCommand()

                                    'cmd = MConExcel.CreateCommand()
                                    cmd.CommandText = "SELECT * "
                                    cmd.CommandText &= " FROM `" & TableName & "`;"

                                    Try

                                        Using reader As OleDbDataReader = cmd.ExecuteReader()

                                            'reader = cmd.ExecuteReader()
                                            dt.Load(reader)

                                            TableName = TableName.Replace("$", "")
                                            dt.TableName = TableName

                                            If dt.Rows.Count > 0 Then

                                                '2022-07-14 15:09, By Cucun, Karena tanggal 13 diubah jadi Non HDR, nama kolom nya jadi F1 F2 F3 dst semua
                                                'Ambil baris pertama sebelum di-delete
                                                'DicColumnName = New Dictionary(Of Integer, String)
                                                For i As Integer = 0 To dt.Columns.Count - 1

                                                    'DicColumnName.Add(i, dt.Columns(i).ColumnName)

                                                    NamaKolom = "" & dt.Rows(0).Item(i)

                                                    If NamaKolom <> "" Then
                                                        dt.Columns(i).ColumnName = NamaKolom
                                                    End If

                                                Next

                                                '2022-07-13 18:02, By Cucun, Karena tidak pakai connection string HDR=Y
                                                'Maka harus buang baris pertama manual
                                                'Kalau pakai HDR=Y, ada kasus 8 baris pertama angka, kemudian sisa nya bukan angka, maka jadi kosong
                                                'Dikasih IMEX=1 percuma
                                                dt.Rows(0).Delete()
                                                'Harus pakai AcceptChanges, jika tidak akan error Deleted row information cannot be accessed through the row
                                                dt.AcceptChanges()

                                            End If

                                            hasil.Tables.Add(dt)

                                        End Using

                                    Catch ex As Exception
                                    End Try

                                End Using

                            End If

                        End Using

                    Next

                    'DicColumnName = Nothing

                End If

            End Using

            Return hasil

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return Nothing

        Finally

            Try
                MConExcel.Close()
            Catch ex As Exception
            End Try

            Try
                MConExcel.Dispose()
            Catch ex As Exception
            End Try

            hasil = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function ReadFileExcel(ByVal FilePath As String, ByVal Columns() As String, ByRef ErrorMessage As String) As String

        Dim hasil As String = ""

        Dim MConExcel As New OleDbConnection

        Try

            Dim QuerySelect As String = ""
            Dim QueryWhere As String = ""

            Dim index As Integer = 0

            For Each Column As String In Columns

                If index = 0 Then
                    QueryWhere &= " AND (TRIM(`" & Column & "`) <> ''"
                Else
                    QuerySelect &= ","
                    QueryWhere &= " OR TRIM(`" & Column & "`) <> ''"
                End If

                QuerySelect &= " IIF(`" & Column & "` <> '', REPLACE(`" & Column & "`,',',''), '') as `" & Column & "`"

                index = index + 1

            Next

            If QueryWhere <> "" Then
                QueryWhere &= ")"
            End If

            MConExcel = ObjCon.SetConnExcel(FilePath)
            MConExcel.Open()

            Dim SB As New StringBuilder
            'Dim dt As DataTable = New DataTable

            Using dtSheets As DataTable = MConExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)

                'Dim dtSheets As DataTable = New DataTable
                'dtSheets = MConExcel.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)

                For Each dr As DataRow In dtSheets.Rows

                    Using dt As DataTable = New DataTable

                        If dr("TABLE_NAME").ToString.EndsWith("$") Or dr("TABLE_NAME").ToString.EndsWith("$'") Then

                            Using cmd As OleDbCommand = MConExcel.CreateCommand()

                                'Dim cmd As OleDbCommand = MConExcel.CreateCommand()
                                cmd.CommandText = "SELECT "
                                cmd.CommandText &= QuerySelect
                                cmd.CommandText &= " FROM `" & dr("TABLE_NAME") & "` WHERE TRUE"
                                cmd.CommandText &= QueryWhere

                                Try

                                    Using reader As OleDbDataReader = cmd.ExecuteReader()

                                        'Dim reader As OleDbDataReader = cmd.ExecuteReader()
                                        dt.Load(reader)

                                        index = 0
                                        For Each dtRow As DataRow In dt.Rows

                                            If index > 0 Then 'Skip baris pertama

                                                If index > 1 Then
                                                    SB.Append("|")
                                                End If

                                                For i As Integer = 0 To Columns.Length - 1
                                                    If i > 0 Then
                                                        SB.Append(",")
                                                    End If
                                                    SB.Append(("" & dtRow.Item(Columns(i)).ToString.Replace("|", "").Replace(",", "")))
                                                Next

                                            End If

                                            index = index + 1

                                        Next

                                    End Using

                                Catch ex As Exception
                                End Try

                            End Using

                        End If

                    End Using

                Next

            End Using

            hasil = SB.ToString

            SB = Nothing

            Return hasil

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return hasil

        Finally

            Try
                MConExcel.Close()
            Catch ex As Exception
            End Try

            Try
                MConExcel.Dispose()
            Catch ex As Exception
            End Try

            GC.Collect()

        End Try

    End Function

    Public Function UploadFile(ByRef UploadedPath As String, ByVal MyFileUpload As FileUpload, ByVal UserCode As String, ByVal PageName As String, ByRef ErrorMessage As String) As Boolean

        Return UploadFile(UploadedPath, MyFileUpload, UserCode, "PathUploadTemp", "", PageName, ErrorMessage)

    End Function

    Public Function UploadFile(ByRef UploadedPath As String, ByVal MyFileUpload As FileUpload, ByVal UserCode As String, ByVal PathUploadConfigName As String, ByVal PageName As String, ByRef ErrorMessage As String) As Boolean

        Return UploadFile(UploadedPath, MyFileUpload, UserCode, PathUploadConfigName, "", PageName, ErrorMessage)

    End Function

    Public Function UploadFile(ByRef UploadedPath As String, ByVal MyFileUpload As FileUpload, ByVal UserCode As String, ByVal PathUploadConfigName As String, ByVal SaveAsFileName As String, ByVal PageName As String, ByRef ErrorMessage As String) As Boolean

        Try

            Dim ServerPath As String = ReadWebConfig(PathUploadConfigName)

            If ServerPath <> "" Then

                Dim AutoNum As String = Date.Now.ToString("yyMMddHHmmssff")

                Dim FileName As String = Path.GetFileName(MyFileUpload.FileName)
                Dim FileNameClear As String = Path.GetFileNameWithoutExtension(FileName)
                Dim ext As String = Path.GetExtension(FileName)

                If SaveAsFileName <> "" Then
                    FileName = SaveAsFileName & ext
                Else
                    '2021-12-15 10:17, By Cucun, tambah PageName, ada case Perubahan Upload Biaya Kirim, CourierDeliveryCost
                    'Tidak bisa dibedakan, mana file2 yang dari fitur tersebut
                    If PageName <> "" Then
                        FileNameClear = PageName & "_" & FileNameClear
                    End If
                    FileName = FileNameClear & "_" & UserCode & "_" & AutoNum & ext
                End If

                If Not Directory.Exists(ServerPath) Then
                    Directory.CreateDirectory(ServerPath)
                End If

                Dim FilePath As String = ServerPath & "\" & FileName

                If File.Exists(FilePath) Then
                    ErrorMessage = FileName & " sudah ada di server!"
                    Return False
                End If

                FilePath = FilePath.Replace("\\", "\")
                MyFileUpload.SaveAs(FilePath)
                UploadedPath = FilePath

                Return True

            Else

                ErrorMessage = "Gagal mengambil setting " & PathUploadConfigName
                Return False

            End If

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return False

        Finally

            MyFileUpload = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function UploadAndResizePic(ByRef UploadedPath As String, ByVal MyFileUpload As FileUpload, ByVal MaxSize As Integer, ByVal UserCode As String, ByVal PathUploadConfigName As String, ByVal SaveAsFileName As String, ByRef ErrorMessage As String) As Boolean

        Try

            Dim ServerPath As String = ReadWebConfig(PathUploadConfigName)

            If ServerPath <> "" Then

                Dim AutoNum As String = Date.Now.ToString("yyMMddHHmmssff")

                Dim FileNameExt As String = Path.GetFileName(MyFileUpload.PostedFile.FileName)
                Dim FileName As String = Path.GetFileNameWithoutExtension(FileNameExt)
                Dim FileExt As String = Path.GetExtension(FileNameExt)

                If SaveAsFileName <> "" Then
                    FileNameExt = SaveAsFileName & FileExt
                Else
                    FileNameExt = FileName & "_" & UserCode & "_" & AutoNum & FileExt
                End If

                If Not Directory.Exists(ServerPath) Then
                    Directory.CreateDirectory(ServerPath)
                End If

                Dim FilePath As String = ServerPath & "\" & FileNameExt

                If File.Exists(FilePath) Then
                    ErrorMessage = FileNameExt & " sudah ada di server!"
                    Return False
                End If

                FilePath = FilePath.Replace("\\" & FileNameExt, "\" & FileNameExt)

                Using postedImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(MyFileUpload.PostedFile.InputStream)

                    Dim OldWidth As Integer = postedImage.Width
                    Dim OldHeight As Integer = postedImage.Height

                    Dim NewValue As Integer = MaxSize

                    Dim Scale As Double = 0

                    Dim OldValue As Integer = 0

                    If OldHeight >= OldWidth Then
                        OldValue = OldHeight
                    Else
                        OldValue = OldWidth
                    End If

                    If OldValue > NewValue Then
                        Scale = NewValue / OldValue
                    Else
                        Scale = 1
                    End If

                    If Scale <> 1 Then

                        Dim NewWidth As Integer = postedImage.Width * Scale
                        Dim NewHeight As Integer = postedImage.Height * Scale

                        Using NewImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(NewWidth, NewHeight)

                            'Dim NewImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(NewWidth, NewHeight)

                            Using thumbGraph As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(NewImage)

                                'Dim thumbGraph As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(NewImage)

                                'Dim ImageRectangle As System.Drawing.Rectangle = New System.Drawing.Rectangle(0, 0, NewWidth, NewHeight)

                                thumbGraph.CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
                                thumbGraph.SmoothingMode = Drawing.Drawing2D.SmoothingMode.HighQuality
                                thumbGraph.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                                thumbGraph.DrawImage(postedImage, New System.Drawing.Rectangle(0, 0, NewWidth, NewHeight))
                                NewImage.Save(FilePath, postedImage.RawFormat)

                            End Using

                        End Using

                    Else

                        MyFileUpload.SaveAs(FilePath)

                    End If

                End Using

                UploadedPath = FilePath

                Return True

            Else

                ErrorMessage = "Gagal mengambil setting " & PathUploadConfigName
                Return False

            End If

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return False

        Finally

            MyFileUpload = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function WriteFileExcel(ByVal data As DataTable, ByVal FolderName As String, ByVal FileName As String, ByRef PesanError As String) As String

        Dim str1 As String = "WriteFileExcel " & FileName

        WriteTracelogTxt(str1 & ", Masuk")

        Dim Result As String = "0|Failed"
        WriteTracelogTxt(str1 & ", Path = ")
        Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\GeneratedReport\" & FolderName & "\")
        WriteTracelogTxt(str1 & ", Path = " & Path)
        Path = Path.Replace("\\", "\")
        WriteTracelogTxt(str1 & ", Path After Replace = " & Path)

        Dim UserDirectory As Boolean = False
        Try
            If Directory.Exists(Path) = True Then
                UserDirectory = True
            Else
                Directory.CreateDirectory(Path)
                UserDirectory = True
            End If
        Catch ex As Exception
            WriteTracelogTxt(str1 & ", User Directory " & Path & " Error : " & ex.Message)
            PesanError = ex.Message
        End Try

        If UserDirectory Then

            WriteTracelogTxt(str1 & ", UserDirectory True")

            Dim FullPath As String = Path & FileName
            FullPath = FullPath.Replace("\\", "\")

            WriteTracelogTxt(str1 & ", UserDirectory FullPath = " & FullPath)

            Dim MConExcel As New OleDbConnection

            WriteTracelogTxt(str1 & ", Declare MConExcel Done")

            Try

                WriteTracelogTxt(str1 & ", Create Column Start")

                Dim index As Integer = 0
                Dim SB As New StringBuilder
                Dim SB2 As New StringBuilder
                Dim invalid_char = New String() {"<BR>", "@", "!", "#", "$", "<B>", "<I>"}
                Dim length_invalid_char As Integer = invalid_char.Length()

                For Each DC As DataColumn In Data.Columns
                    If index > 0 Then
                        SB.Append(",")
                        SB2.Append(",")
                    End If
                    SB.Append("`" & DC.ColumnName.ToUpper & "` VARCHAR(255)")
                    SB2.Append("`" & DC.ColumnName.ToUpper & "`")

                    For a = 0 To length_invalid_char - 1
                        SB.Replace(invalid_char(a), " ")
                        SB2.Replace(invalid_char(a), " ")
                    Next

                    index = index + 1
                Next

                WriteTracelogTxt(str1 & ", Create Column End")

                Dim Kolom As String = SB.ToString

                WriteTracelogTxt(str1 & ", Kolom = SB.ToString")

                MConExcel = ObjCon.SetConnExcelWrite(FullPath)
                WriteTracelogTxt(str1 & ", MConExcel Set to FullPath " & FullPath)
                MConExcel.Open()

                WriteTracelogTxt(str1 & ", MConExcel Open Done")

                Using cmd As OleDbCommand = MConExcel.CreateCommand()

                    cmd.CommandText = "CREATE TABLE `Sheet1` (" & Kolom & ")"
                    cmd.ExecuteNonQuery()

                    WriteTracelogTxt(str1 & ", MConExcel CREATE TABLE `Sheet1` Done")

                    Dim SB3 As New StringBuilder
                    Dim SqlParam As New Dictionary(Of String, String)

                    Dim CurrRow As Integer = 0
                    Dim MaxRow As Integer = 1 'OleDb Excel ga bisa insert lebih dari 1 baris sekaligus
                    Dim MaxDataRow As Integer = data.Rows.Count - 1

                    For i As Integer = 0 To data.Rows.Count - 1

                        If CurrRow = 0 Then

                            SB3 = New StringBuilder
                            SB3.Append("INSERT INTO `Sheet1` (" & SB2.ToString & ") VALUES ")

                            cmd.Parameters.Clear()
                            SqlParam = New Dictionary(Of String, String)

                        Else
                            SB3.Append(",")
                        End If

                        SB3.Append("(")
                        For j As Integer = 0 To data.Columns.Count - 1

                            If j > 0 Then
                                SB3.Append(",").Replace("<br>", " ")
                            End If

                            SB3.Append("@col" & i & j)
                            SqlParam.Add("@col" & i & j, "" & data.Rows(i).Item(j))

                        Next
                        SB3.Append(")")

                        CurrRow += 1

                        If CurrRow >= MaxRow Or i = MaxDataRow Then

                            cmd.CommandText = SB3.ToString

                            If Not IsNothing(SqlParam) Then

                                For Each Parameter As KeyValuePair(Of String, String) In SqlParam

                                    cmd.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                                Next

                            End If

                            cmd.ExecuteNonQuery()

                            CurrRow = 0

                        End If

                    Next

                    SqlParam = Nothing
                    SB3 = Nothing

                    WriteTracelogTxt(str1 & ", MConExcel INSERT INTO `Sheet1` Done")

                End Using

                Result = "1|" & FullPath

            Catch ex As Exception

                WriteTracelogTxt(str1 & ", " & ex.Message)
                PesanError = ex.Message

            Finally

                Try
                    MConExcel.Close()
                Catch ex As Exception
                End Try

                Try
                    MConExcel.Dispose()
                Catch ex As Exception
                End Try

                GC.Collect()

            End Try

        Else

            PesanError = "User Directory " & Path & " NOT Exists"

        End If

        Return Result

    End Function

    Public Function WriteFileExcelMultiSheet(ByVal data As DataSet, ByVal UseSheetsName As Boolean, ByVal SheetsName() As String, ByVal FolderName As String, ByVal FileName As String, ByRef PesanError As String) As String

        Return WriteFileExcelMultiSheet(data, UseSheetsName, SheetsName, False, FolderName, FileName, PesanError)

    End Function

    Public Function WriteFileExcelMultiSheet(ByVal data As DataSet, ByVal UseSheetsName As Boolean, ByVal SheetsName() As String, ByVal FirstRowAsHeader As Boolean, ByVal FolderName As String, ByVal FileName As String, ByRef PesanError As String) As String

        Dim Result As String = "0|Failed"
        Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\GeneratedReport\" & FolderName & "\")
        Path = Path.Replace("\\", "\")

        Dim UserDirectory As Boolean = False
        Try
            If Directory.Exists(Path) = True Then
                UserDirectory = True
            Else
                Directory.CreateDirectory(Path)
                UserDirectory = True
            End If
        Catch ex As Exception
            WriteTracelogTxt(ex.Message)
            PesanError = ex.Message
        End Try

        If UserDirectory Then

            Dim FullPath As String = Path & FileName
            FullPath = FullPath.Replace("\\", "\")

            Dim MConExcel As New OleDbConnection

            Try

                Dim index As Integer = 0
                Dim SheetNum As Integer = 1

                Dim SB As StringBuilder
                Dim SB2 As StringBuilder
                Dim SB3 As StringBuilder

                Dim SqlParam As New Dictionary(Of String, String)

                Dim SheetName As String = ""
                Dim Kolom As String = ""

                If data.Tables.Count > 0 Then
                    MConExcel = ObjCon.SetConnExcelWrite(FullPath)
                End If

                For Each tbl As DataTable In data.Tables

                    SB = New StringBuilder
                    SB2 = New StringBuilder

                    If Not FirstRowAsHeader Then
                        index = 0
                        For Each DC As DataColumn In tbl.Columns
                            If index > 0 Then
                                SB.Append(",")
                                SB2.Append(",")
                            End If
                            SB.Append("`" & DC.ColumnName.ToUpper & "` VARCHAR(255)")
                            SB2.Append("`" & DC.ColumnName.ToUpper & "`")
                            index = index + 1
                        Next
                    Else

                        If tbl.Rows.Count > 0 Then

                            index = 0
                            For Each DC As DataColumn In tbl.Columns
                                If index > 0 Then
                                    SB.Append(",")
                                    SB2.Append(",")
                                End If
                                SB.Append("`" & tbl.Rows(0).Item(index) & "` VARCHAR(255)")
                                SB2.Append("`" & tbl.Rows(0).Item(index) & "`")
                                index = index + 1
                            Next

                        End If

                    End If

                    Kolom = SB.ToString

                    If MConExcel.State <> ConnectionState.Open Then
                        MConExcel.Open()
                    End If

                    Try
                        If UseSheetsName Then
                            SheetName = "`" & SheetsName(SheetNum - 1) & "`"
                        Else
                            SheetName = "`" & tbl.TableName & "`"
                        End If
                    Catch
                        SheetName = "`Sheet" & SheetNum & "`"
                    End Try

                    Using cmd As OleDbCommand = MConExcel.CreateCommand()

                        cmd.CommandText = "CREATE TABLE " & SheetName & " (" & Kolom & ")"
                        cmd.ExecuteNonQuery()

                        SB3 = New StringBuilder
                        SqlParam = New Dictionary(Of String, String)

                        Dim CurrRow As Integer = 0
                        Dim MaxRow As Integer = 1 'OleDb Excel ga bisa insert lebih dari 1 baris sekaligus
                        Dim MaxDataRow As Integer = tbl.Rows.Count - 1

                        For i As Integer = 0 To tbl.Rows.Count - 1

                            If FirstRowAsHeader And i = 0 Then
                                'Jika First Row Sebagai Header dan sekarang row pertama, maka skip jangan tulis apa2
                            Else

                                If CurrRow = 0 Then

                                    SB3 = New StringBuilder
                                    SB3.Append("INSERT INTO " & SheetName & " (" & SB2.ToString & ") VALUES ")

                                    cmd.Parameters.Clear()
                                    SqlParam = New Dictionary(Of String, String)

                                Else
                                    SB3.Append(",")
                                End If

                                SB3.Append("(")
                                For j As Integer = 0 To tbl.Columns.Count - 1

                                    If j > 0 Then
                                        SB3.Append(",")
                                    End If

                                    SB3.Append("@col" & i & j)
                                    SqlParam.Add("@col" & i & j, "" & tbl.Rows(i).Item(j))

                                Next
                                SB3.Append(")")

                                CurrRow += 1

                                If CurrRow >= MaxRow Or i = MaxDataRow Then

                                    cmd.CommandText = SB3.ToString

                                    If Not IsNothing(SqlParam) Then

                                        For Each Parameter As KeyValuePair(Of String, String) In SqlParam

                                            cmd.Parameters.AddWithValue(Parameter.Key, Parameter.Value)

                                        Next

                                    End If

                                    cmd.ExecuteNonQuery()

                                    CurrRow = 0

                                End If

                            End If

                        Next

                    End Using

                    SheetNum = SheetNum + 1

                Next

                SB = Nothing
                SB2 = Nothing
                SB3 = Nothing

                SqlParam = Nothing

                Result = "1|" & FullPath

                Return Result

            Catch ex As Exception

                WriteTracelogTxt(ex.Message)
                PesanError = ex.Message

                Return Result

            Finally

                Try
                    MConExcel.Close()
                Catch ex As Exception
                End Try

                Try
                    MConExcel.Dispose()
                Catch ex As Exception
                End Try

                GC.Collect()

            End Try

        Else

            PesanError = "User Directory " & Path & " NOT Exists"

            Return Result

        End If

    End Function

#End Region

#Region "Tools"

    Public Function isChecked_CheckBoxList(ByVal chkList As CheckBoxList) As String()

        Dim hasil(1) As String

        Try
            Dim count As Integer = 0

            For Each li As ListItem In chkList.Items

                If li.Selected = True Then
                    count = count + 1
                End If

            Next

            If count = chkList.Items.Count Then
                hasil(0) = "2" 'kecentang semua
            ElseIf count = 0 Then
                hasil(0) = "0" 'ga ad yang kecentang
            Else
                hasil(0) = "1" 'kecentang beberapa
            End If
            hasil(1) = ""

            Return hasil

        Catch ex As Exception

            hasil(0) = "9"
            hasil(1) = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return hasil

        End Try

    End Function

#End Region

#Region "GenerateReport"

    Public Function GenerateReport(ByVal ReportName As String, ByVal User As String, ByVal AutoNum As String, ByVal Data As String) As String

        Return GenerateReport(ReportName, ".csv", "GeneratedReport", User, AutoNum, Data)

    End Function

    Public Function GenerateReport(ByVal ReportName As String, ByVal FolderName As String, ByVal User As String, ByVal AutoNum As String, ByVal Data As String) As String

        Return GenerateReport(ReportName, ".csv", FolderName, User, AutoNum, Data)

    End Function

    Public Function GenerateReport(ByVal ReportName As String, ByVal ExtensionFile As String, ByVal FolderName As String, ByVal User As String, ByVal AutoNum As String, ByVal Data As String) As String

        Dim File As String = ""
        Dim Hasil As String = "0||"

        If FolderName.Trim = "" Then
            FolderName = "GeneratedReport"
        End If

        Dim PhysicalWebPath As String = System.Web.HttpContext.Current.Server.MapPath("~\" & FolderName)
        PhysicalWebPath = PhysicalWebPath.Replace("\\" & FolderName, "\" & FolderName)

        Dim PhysicalUserDirectory As String = PhysicalWebPath & "\" & User
        PhysicalUserDirectory = PhysicalUserDirectory.Replace("\\" & User, "\" & User)

        Dim WebUserDirectory As String = FolderName & "\" & User
        WebUserDirectory = WebUserDirectory.Replace("\\" & User, "\" & User)

        Dim UserDirectoryIsExists As Boolean = False

        Try
            If Directory.Exists(PhysicalUserDirectory) Then
                UserDirectoryIsExists = True
            Else
                Directory.CreateDirectory(PhysicalUserDirectory)
                UserDirectoryIsExists = True
            End If
        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            UserDirectoryIsExists = False

        End Try

        Try

            If UserDirectoryIsExists Then

                Dim FileWebPath As String = ""

                If AutoNum <> "" Then
                    AutoNum = "_" & AutoNum
                End If

                ExtensionFile = ExtensionFile.Trim
                If ExtensionFile <> "" Then
                    While ExtensionFile.StartsWith(".")
                        ExtensionFile = ExtensionFile.Substring(1, ExtensionFile.Length - 1)
                    End While
                    ExtensionFile = "." & ExtensionFile
                End If
                ExtensionFile = ExtensionFile.ToLower

                Dim ReportFullName As String = ReportName & AutoNum & ExtensionFile

                FileWebPath = WebUserDirectory & "\" & ReportFullName
                FileWebPath = FileWebPath.Replace("\\" & ReportFullName, "\" & ReportFullName)

                File = PhysicalUserDirectory & "\" & ReportFullName
                File = File.Replace("\\" & ReportFullName, "\" & ReportFullName)

                Using SW As StreamWriter = New StreamWriter(File)
                    SW.Write(Data)
                    SW.Flush()
                End Using

                Hasil = "1|" & File & "|" & FileWebPath

            End If
        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Gagal Tulis File ke : " & File & vbCrLf & " Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Hasil = "0|" & File & "|"

        Finally

        End Try

        Return Hasil

    End Function

#End Region

#Region "Fungsi Multi Select"

    Public Function SetDataMulti(ByVal btn As Button, ByVal chkList As CheckBoxList, Optional ByVal isSetTool As Boolean = True) As String

        Try

            Dim hasil() As String = isChecked_CheckBoxList(chkList)
            Dim hasil1 As String = ""

            If hasil(0) = "9" Then

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & hasil(1)
                WriteTracelogTxt(Pesan)

            Else
                If hasil(0) = "0" Or hasil(0) = "2" Then
                    hasil1 = ""
                    If isSetTool Then
                        Dim ClassCss As String = ConfigurationManager.AppSettings("CssDefaultButton")
                        btn.Attributes.Remove("class")
                        btn.Text = "..."
                        btn.CssClass = ClassCss
                    End If
                Else
                    hasil1 = GetMulti(chkList)
                    If isSetTool Then
                        Dim ClassCss As String = ConfigurationManager.AppSettings("CssMultiSelectButton")
                        btn.Attributes.Remove("class")
                        btn.Text = "v"
                        btn.CssClass = ClassCss
                    End If
                End If
            End If

            Return hasil1

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Function SetDataMultiName(ByVal btn As Button, ByVal chkList As CheckBoxList, Optional ByVal isSetTool As Boolean = True) As String

        Try

            Dim hasil() As String = isChecked_CheckBoxList(chkList)
            Dim hasil1 As String = ""

            If hasil(0) = "9" Then

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & hasil(1)
                WriteTracelogTxt(Pesan)

            Else
                If hasil(0) = "0" Or hasil(0) = "2" Then
                    hasil1 = ""
                    If isSetTool Then
                        Dim ClassCss As String = ConfigurationManager.AppSettings("CssDefaultButton")
                        btn.Attributes.Remove("class")
                        btn.Text = "..."
                        btn.CssClass = ClassCss
                    End If
                Else
                    hasil1 = GetMultiName(chkList)
                    If isSetTool Then
                        Dim ClassCss As String = ConfigurationManager.AppSettings("CssMultiSelectButton")
                        btn.Attributes.Remove("class")
                        btn.Text = "v"
                        btn.CssClass = ClassCss
                    End If
                End If
            End If

            Return hasil1

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

    End Function

    Public Sub SelectAll(ByVal chkList As CheckBoxList)

        For Each li As ListItem In chkList.Items
            li.Selected = True
        Next

    End Sub

    Public Sub DeSelectAll(ByVal chkList As CheckBoxList)

        For Each li As ListItem In chkList.Items
            li.Selected = False
        Next

    End Sub

    Public Function GetMulti(ByVal chkList As CheckBoxList) As String

        Dim hasil As String = ""

        Try

            For Each li As ListItem In chkList.Items

                If li.Selected Then

                    If hasil <> "" Then
                        hasil &= "|"
                    End If
                    hasil &= li.Value

                End If

            Next

            Return hasil

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

        Return hasil

    End Function

    Public Function GetMultiName(ByVal chkList As CheckBoxList) As String

        Dim hasil As String = ""

        Try

            For Each li As ListItem In chkList.Items

                If li.Selected Then

                    If hasil <> "" Then
                        hasil &= " / "
                    End If
                    hasil &= li.Text

                End If

            Next

            Return hasil

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        End Try

        Return hasil

    End Function

    Public Sub SetMulti(ByVal chkList As CheckBoxList, ByVal ListData As String)

        Try
            If IsNothing(ListData) Then
                DeSelectAll(chkList)
            Else
                For Each li As ListItem In chkList.Items

                    Dim i As String = "|" & li.Value & "|"
                    Dim j As String = "|" & ListData & "|"

                    If j.Contains(i) Then
                        li.Selected = True
                    Else
                        li.Selected = False
                    End If

                Next
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

    End Sub

#End Region

    Public Function AccountIsHO(ByVal MyCode As String) As Boolean

        If MyCode.ToUpper.StartsWith("HO") Or MyCode.ToUpper.StartsWith("FA") Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function GetPerformBulananTipeHitung() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            dt = ConvertStringToDatatable("Value,Display|,|awb,JUMLAH AWB|weight,TOTAL BERAT|deliverycost,RUPIAH")

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function ConvertStringToDatatableForSJ(ByVal Param As String) As DataTable

        Dim dtResult As DataTable = Nothing

        Try

            Dim Result As String() = Param.ToString.Split("|") 'pemisah row

            For i As Integer = 0 To Result.Length - 1

                If i = 0 Then 'column header
                    If dtResult Is Nothing Then
                        dtResult = New DataTable
                    End If
                Else 'data
                    dtResult.Rows.Add()
                End If

                Dim Result2Split As String() = Result(i).Split(",") 'pemisah kolom

                For j As Integer = 0 To Result2Split.Length - 1
                    If i = 0 Then 'column header
                        If Result2Split(j) = "Weight" Then
                            dtResult.Columns.Add(Result2Split(j), GetType(Double))
                        ElseIf Result2Split(j) = "Paket" Then
                            dtResult.Columns.Add(Result2Split(j), GetType(Integer))
                        Else
                            dtResult.Columns.Add(Result2Split(j))
                        End If
                    Else 'data
                        dtResult.Rows(i - 1).Item(j) = Result2Split(j)
                    End If
                Next

            Next

            Return dtResult

        Catch ex As Exception

            Return dtResult

        Finally

            dtResult = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function EmailValidasi(ByVal Email As String, ByRef Pesan As String, Optional ByVal AllowEmpty As Boolean = False) As Boolean

        Dim Result As Boolean = False

        If Email.Trim = "" Then
            If AllowEmpty = False Then
                Pesan = "tidak boleh kosong!"
                Return False
            End If
        End If

        Pesan = "tidak valid!"

        If Email.Contains("@") = False Or Email.Contains(".") = False Then
            GoTo Skip
        End If

        If Email.Contains(" ") Then
            Pesan = "tidak boleh ada spasi!"
            GoTo Skip
        End If

        If Email.LastIndexOf("@") > Email.LastIndexOf(".") Then
            GoTo Skip
        End If

        If Email.Contains("..") Or Email.Contains("@@") _
        Or Email.Contains(".@") Or Email.Contains("@.") _
        Or Email.Contains("-@") Or Email.Contains("@-") Then
            GoTo Skip
        End If

        If Email.StartsWith(".") Or Email.EndsWith(".") _
        Or Email.StartsWith("-") Or Email.EndsWith("-") Then
            GoTo Skip
        End If

        If Email.IndexOfAny(New Char() {"!", "#", "$", "%", "^", "&", "*", ";", ":", "'", """", "?", ","}) >= 0 Then
            GoTo Skip
        End If

        If Email.IndexOf("@") <> Email.LastIndexOf("@") Then
            GoTo Skip
        End If


        Pesan = ""
        Result = True

Skip:

        Return Result
    End Function

    Public Function ProsesZip(ByVal FilesToZip As String(), ByVal OutputFolder As String, ByVal OutputFilename As String, ByRef PesanError As String) As String

        Dim Hasil As String = ""

        Try

            Dim PhysicalWebPath As String = System.Web.HttpContext.Current.Server.MapPath("~\" & OutputFolder)
            Dim UserDirectoryIsExists As Boolean = False

            Try
                If Directory.Exists(PhysicalWebPath) = True Then
                    UserDirectoryIsExists = True
                Else
                    Directory.CreateDirectory(PhysicalWebPath)
                    UserDirectoryIsExists = True
                End If
            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                UserDirectoryIsExists = False

            End Try

            If UserDirectoryIsExists Then

                Dim outPathname As String = PhysicalWebPath & "\" & OutputFilename
                outPathname = outPathname.Replace("\\" & OutputFilename, "\" & OutputFilename)

                If CreateZipSelectedFiles(outPathname, "", FilesToZip, PesanError) Then

                    Hasil = outPathname

                End If

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return Hasil

    End Function

    Private Function CreateZipSelectedFiles(ByVal outPathname As String, ByVal password As String, ByVal filenames As String(), ByRef PesanError As String) As Boolean

        Try

            Using fsOut As FileStream = File.Create(outPathname)

                'Dim fsOut As FileStream = File.Create(outPathname)

                Using zipStream As ZipOutputStream = New ZipOutputStream(fsOut)

                    'Dim zipStream As ZipOutputStream = New ZipOutputStream(fsOut)

                    zipStream.SetLevel(3) '0-9, 9 being the highest level of compression

                    zipStream.Password = password 'optional. Null is the same as not setting. Required if using AES.

                    ' This setting will strip the leading part of the folder path in the entries, to
                    ' make the entries relative to the starting folder.
                    ' To include the full path for each entry up to the drive root, assign folderOffset = 0.
                    ' int folderOffset = folderName.Length + (folderName.EndsWith("\\") ? 0 : 1);
                    ' Dim folderOffset As Integer = folderName.Length + (If(folderName.EndsWith("\"), 0, 1))

                    CompressSelectedFiles(filenames, zipStream)

                    zipStream.IsStreamOwner = True 'Makes the Close also Close the underlying stream
                    zipStream.Close()

                End Using

                fsOut.Close()

            End Using

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    Private Sub CompressSelectedFiles(ByVal filenames As String(), ByVal zipStream As ZipOutputStream)

        Try

            For Each filename As String In filenames

                Dim fi As FileInfo = New FileInfo(filename)

                Dim entryName As String = filename
                entryName = fi.Name

                Dim newEntry As ZipEntry = New ZipEntry(entryName)
                newEntry.DateTime = fi.LastWriteTime 'Note the zip format stores 2 second granularity

                ' Specifying the AESKeySize triggers AES encryption. Allowable values are 0 (off), 128 or 256.
                ' A password on the ZipOutputStream is required if using AES.
                '   newEntry.AESKeySize = 256;

                ' To permit the zip to be unpacked by built-in extractor in WinXP and Server2003, WinZip 8, Java, and other older code,
                ' you need to do one of the following: Specify UseZip64.Off, or set the Size.
                ' If the file may be bigger than 4GB, or you do not need WinXP built-in compatibility, you do not need either,
                ' but the zip will be in Zip64 format which not all utilities can understand.
                '   zipStream.UseZip64 = UseZip64.Off;
                newEntry.Size = fi.Length

                zipStream.PutNextEntry(newEntry)

                ' Zip the file in buffered chunks
                ' the "using" will close the stream even if an exception occurs
                Dim buffer As Byte() = New Byte(4096) {}
                Using streamReader As FileStream = File.OpenRead(filename)
                    StreamUtils.Copy(streamReader, zipStream, buffer)
                    streamReader.Close()
                End Using

                zipStream.CloseEntry()

            Next

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

    End Sub

    Public Function ReadCSVFile(ByVal FilePath As String, ByVal SkipFirstRow As Boolean, ByRef PesanError As String) As String

        Dim hasil As String = ""
        Dim SB As New StringBuilder

        Try

            Dim line As String
            Dim index As Integer = 0
            Dim indexW As Integer = 0

            Using SR As New StreamReader(FilePath)

                While Not SR.EndOfStream

                    line = SR.ReadLine

                    If indexW > 0 Then

                        SB.Append("|")

                    End If

                    If SkipFirstRow And index = 0 Then
                        'Do Nothing
                    Else

                        SB.Append(line)
                        indexW += 1

                    End If

                    index += 1

                End While

            End Using

            hasil = SB.ToString

        Catch ex As Exception

            PesanError = ex.Message
            hasil = ""

        Finally

            SB = Nothing
            GC.Collect()

        End Try

        Return hasil

    End Function

    Public Function GroupData(ByVal Data As String, ByVal Separator() As String, ByVal MaxDataPerRow As Integer, ByRef ErrorMessage As String) As String()

        Dim Hasil() As String = Nothing

        Try

            Dim MyData() As String = Data.Split(Separator, StringSplitOptions.None)

            If MyData.Length > 0 Then

                Dim ResRow As Integer = Math.Ceiling(MyData.Length / MaxDataPerRow)

                ReDim Hasil(ResRow)

                Dim CurrRowHasil As Integer = 0
                Dim DataLength As Integer = MyData.Length - 1

                Dim CurrRowData As Integer = 0
                Dim SB As New StringBuilder

                Dim AllSeparator As String = ""

                For Each TempSeparator As String In Separator

                    AllSeparator &= TempSeparator

                Next

                For i As Integer = 0 To DataLength

                    If i > 0 Then
                        SB.Append(AllSeparator)
                    End If

                    SB.Append(MyData(i))

                    CurrRowData += 1

                    If CurrRowData >= MaxDataPerRow Or i = DataLength Then
                        Hasil(CurrRowHasil) = SB.ToString
                        CurrRowHasil += 1
                        CurrRowData = 0
                    End If

                Next

            End If

        Catch ex As Exception

            ErrorMessage = ex.Message

        End Try

        Return Hasil

    End Function

    Public Function EmailSend(ByVal _MailTo As String, ByVal _MailCc As String, ByVal _Subject As String, ByVal _Body As String, ByVal _Filename As String _
    , ByVal eType As String, Optional ByVal HandlingLabel As String = "", Optional ByVal SleepTime As Integer = 150 _
    , Optional ByVal BccIndopaket As Boolean = False, Optional ByVal _Filenames As String() = Nothing, Optional ByVal _MailReplyTo As String = "") As Boolean 'kirim email

        'Dim _EmailFrom As String = "budil@indomaret.co.id"
        'Dim _EmailUsername As String = "budil@indomaret.co.id"
        'Dim _EmailPassword As String = "budil@idm16"
        'Dim _EmailServer As String = "172.20.20.240" 'Lama 192.168.2.240
        'Dim _EmailPort As Integer = 25
        'Dim _EmailDisplayName As String = "budil@indomaret.co.id"
        'Dim _EmailReplyTo As String = ""

        Dim _EmailFrom As String = "support_it@indopaket.co.id"
        Dim _EmailUsername As String = "support_it@indopaket.co.id"
        Dim _EmailPassword As String = "it@ipp20"
        'Dim _EmailServer As String = "172.20.20.85"
        Dim _EmailServer As String = "mail1.indopaket.co.id"
        Dim _EmailPort As Integer = 25
        Dim _EmailDisplayName As String = "INDOPAKET"
        Dim _EmailReplyTo As String = ""

        Try

            If _MailTo.Trim <> "" Then

                '=== Setup Email
                Dim ConfigEmailServer As String = ""
                Try
                    ConfigEmailServer = "" & ConfigurationManager.AppSettings("EmailServerIndopaket")
                Catch ex As Exception
                    ConfigEmailServer = ""
                End Try

                If ConfigEmailServer = "" Then
                    'ConfigEmailServer = "172.20.20.85"
                    ConfigEmailServer = "mail1.indopaket.co.id"
                End If

                Dim TracelogEmail As String = ""
                Try
                    TracelogEmail = "" & ConfigurationManager.AppSettings("TracelogEmail")
                Catch ex As Exception
                    TracelogEmail = ""
                End Try

                If _MailReplyTo <> "" Then
                    _EmailReplyTo = _MailReplyTo
                End If

                If TracelogEmail <> "0" Then

                    Try

                        Dim exMessage As String = "Debug Email Server : " & _EmailServer

                        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                        Dim Pesan As String = ""
                        Pesan = "Dari ClsFungsi, Proses " & methodName & ", " & exMessage
                        WriteTracelogTxt(Pesan)

                    Catch ex As Exception

                        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                        Dim Pesan As String = ""
                        Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & "Debug Email Server : " & _EmailServer & ", Error : " & ex.Message
                        WriteTracelogTxt(Pesan)

                    End Try

                End If

                Dim _EmailTo As String = _MailTo
                Dim _EmailCc As String = _MailCc
                Dim _EmailBcc As String = "budil@indopaket.co.id;cucun@indopaket.co.id;alvinadi.widjaja@indopaket.co.id"
                If BccIndopaket Then
                    _EmailBcc = "budil@indopaket.co.id;cucun@indopaket.co.id;alvinadi.widjaja@indopaket.co.id"
                End If

                Using MyMailMessage As New MailMessage()

                    'Dim MyMailMessage As New MailMessage()

                    MyMailMessage.From = New MailAddress(_EmailFrom, _EmailDisplayName)
                    MyMailMessage.Sender = New MailAddress(_EmailFrom, _EmailDisplayName)

                    Using dtEmailList As DataTable = New DataTable

                        'Dim dtEmailList As New DataTable
                        dtEmailList.Columns.Add("Email")

                        If _EmailTo <> "" Then

                            Dim ListTo As String() = _EmailTo.Split(";")
                            For i As Integer = 0 To ListTo.Length - 1
                                If ListTo(i) <> "" Then
                                    Dim ListTo2 As String() = ListTo(i).Split(",")
                                    For j As Integer = 0 To ListTo2.Length - 1
                                        If ListTo2(j) <> "" Then
                                            dtEmailList.Rows.Add(New String() {ListTo2(j)})
                                        End If
                                    Next
                                End If
                            Next

                            If dtEmailList.Rows.Count > 0 Then
                                For i As Integer = 0 To dtEmailList.Rows.Count - 1
                                    MyMailMessage.To.Add(dtEmailList.Rows(i).Item("Email").ToString)
                                Next
                            End If

                        End If


                        If _EmailCc <> "" Then

                            dtEmailList.Rows.Clear()

                            Dim ListCc As String() = _EmailCc.Split(";")
                            For i As Integer = 0 To ListCc.Length - 1
                                If ListCc(i) <> "" Then
                                    Dim ListCc2 As String() = ListCc(i).Split(",")
                                    For j As Integer = 0 To ListCc2.Length - 1
                                        If ListCc2(j) <> "" Then
                                            dtEmailList.Rows.Add(New String() {ListCc2(j)})
                                        End If
                                    Next
                                End If
                            Next

                            If dtEmailList.Rows.Count > 0 Then
                                For i As Integer = 0 To dtEmailList.Rows.Count - 1
                                    MyMailMessage.CC.Add(dtEmailList.Rows(i).Item("Email").ToString)
                                Next
                            End If

                        End If


                        If _EmailBcc <> "" Then

                            dtEmailList.Rows.Clear()

                            Dim ListBcc As String() = _EmailBcc.Split(";")
                            For i As Integer = 0 To ListBcc.Length - 1
                                If ListBcc(i) <> "" Then
                                    Dim ListBcc2 As String() = ListBcc(i).Split(",")
                                    For j As Integer = 0 To ListBcc2.Length - 1
                                        If ListBcc2(j) <> "" Then
                                            dtEmailList.Rows.Add(New String() {ListBcc2(j)})
                                        End If
                                    Next
                                End If
                            Next

                            If dtEmailList.Rows.Count > 0 Then
                                For i As Integer = 0 To dtEmailList.Rows.Count - 1
                                    MyMailMessage.Bcc.Add(dtEmailList.Rows(i).Item("Email").ToString)
                                Next
                            End If

                        End If

                    End Using 'dtEmailList

                    If _EmailReplyTo <> "" Then
                        Dim ReplyTo As MailAddress = New MailAddress(_EmailReplyTo, _EmailDisplayName)
                        MyMailMessage.ReplyTo = ReplyTo
                    End If

                    MyMailMessage.Subject = _Subject
                    MyMailMessage.Body = _Body
                    MyMailMessage.IsBodyHtml = True

                    'Dim Attach As System.Net.Mail.Attachment = Nothing
                    If _Filename.Trim <> "" Then
                        If IO.File.Exists(_Filename) Then
                            Using Attach As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(_Filename)
                                MyMailMessage.Attachments.Add(Attach)
                            End Using
                        End If
                    End If

                    'If AttachHandlingLabel Then
                    '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\awayheat.jpg"))
                    '    MyMailMessage.Attachments.Add(Attach)
                    '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\fragile.jpg"))
                    '    MyMailMessage.Attachments.Add(Attach)
                    '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\heavy.jpg"))
                    '    MyMailMessage.Attachments.Add(Attach)
                    '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\keepdry.jpg"))
                    '    MyMailMessage.Attachments.Add(Attach)
                    '    Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\thiswayup.jpg"))
                    '    MyMailMessage.Attachments.Add(Attach)
                    'End If

                    'If HandlingLabel <> "" Then
                    '    If HandlingLabel.Contains("HEAT") Then
                    '        Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\awayheat.jpg"))
                    '        MyMailMessage.Attachments.Add(Attach)
                    '    End If
                    '    If HandlingLabel.Contains("FRAGILE") Then
                    '        Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\fragile.jpg"))
                    '        MyMailMessage.Attachments.Add(Attach)
                    '    End If
                    '    If HandlingLabel.Contains("HEAVY") Then
                    '        Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\heavy.jpg"))
                    '        MyMailMessage.Attachments.Add(Attach)
                    '    End If
                    '    If HandlingLabel.Contains("DRY") Then
                    '        Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\keepdry.jpg"))
                    '        MyMailMessage.Attachments.Add(Attach)
                    '    End If
                    '    If HandlingLabel.Contains("UPRIGHT") Then
                    '        Attach = New System.Net.Mail.Attachment(Server.MapPath("Rpt\images\thiswayup.jpg"))
                    '        MyMailMessage.Attachments.Add(Attach)
                    '    End If
                    'End If

                    If _Filenames Is Nothing = False Then
                        For i As Integer = 0 To _Filenames.Length - 1
                            If ("" & _Filenames(i)).Trim <> "" Then
                                If IO.File.Exists(_Filenames(i)) Then
                                    Using Attach As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(_Filenames(i))
                                        MyMailMessage.Attachments.Add(Attach)
                                    End Using
                                End If
                            End If
                        Next
                    End If

                    ''=== Setup POP3
                    'Dim email As Pop3Client
                    'Try
                    '    email = New Pop3Client(_EmailUsername, _EmailPassword, _EmailServer)
                    '    email.OpenInbox()
                    'Catch ex As Exception
                    'End Try


                    '=== Setup SMTP

                    Dim SMTPServer As SmtpClient = New SmtpClient(_EmailServer)

                    SMTPServer.Port = _EmailPort

                    SMTPServer.Credentials = New System.Net.NetworkCredential(_EmailUsername, _EmailPassword)

                    SMTPServer.Send(MyMailMessage)

                    System.Threading.Thread.Sleep(SleepTime)

                    '.Net 2.0 ga ada SmtpClient Dispose
                    'Pengganti nya Set jadi Null lalu panggil GC.Collect()
                    SMTPServer = Nothing
                    GC.Collect()

                    'MyMailMessage = Nothing

                    'Try
                    '    Attach.Dispose()
                    'Catch ex As Exception
                    'End Try

                    'Try
                    '    email.CloseConnection()
                    'Catch ex As Exception
                    'End Try

                End Using 'MyMailMessage

            Else

                Dim exMessage As String = "Skip Email : " & _Subject

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", " & exMessage
                WriteTracelogTxt(Pesan)

            End If

            Return True

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & ", MailFrom : " & _EmailFrom & vbCrLf & "MailTo : " & _MailTo & vbCrLf & "Subject : " & _Subject & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    Public Function DatatableToHTML(ByVal dtData As DataTable) As String

        Dim dString As New StringBuilder

        Try

            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then

                    dString.Append("<table border=""1"" cellpadding=""5"" style=""border-collapse:collapse"">")
                    dString.Append(DatatableToHTMLHeader(dtData))
                    dString.Append(DatatableToHTMLBody(dtData))
                    dString.Append("</table>")

                End If
            End If

            Return ("" & dString.ToString)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        Finally

            dString = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DatatableToHTMLHeader(ByVal dtData As DataTable) As String

        Dim dString As New StringBuilder

        Try

            dString.Append("<thead><tr style=""background-color:Yellow"">")
            For Each dCol As DataColumn In dtData.Columns
                dString.AppendFormat("<th>{0}</th>", dCol.ColumnName)
            Next
            dString.Append("</tr></thead>")

            Return dString.ToString

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        Finally

            dString = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DatatableToHTMLBody(ByVal dtData As DataTable) As String

        Dim dString As New StringBuilder

        Try

            dString.Append("<tbody>")
            For Each dRow As DataRow In dtData.Rows
                Dim IsBlank As Boolean = (("" & dRow(0)).ToString.Trim = "")
                dString.Append("<tr" & IIf(IsBlank, " style=""background-color:Black""", "") & ">")
                For dCount As Integer = 0 To dtData.Columns.Count - 1
                    dString.AppendFormat("<td>{0}</td>", dRow(dCount))
                Next
                dString.Append("</tr>")
            Next
            dString.Append("</tbody>")

            Return dString.ToString

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return ""

        Finally

            dString = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetRegionList(ByVal PageName As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.GetRegionList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetStationList(ByVal PageName As String, ByVal MyStation As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetStationList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                If MyStation <> "" And Not MyStation.ToUpper.StartsWith("HO") Then
                    dt.DefaultView.RowFilter = "Value = '" & MyStation & "'" 'khusus station sendiri
                End If

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetHubList(ByVal PageName As String, ByVal MyHub As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            dt = GetHubListNoFilter()

            If dt.TableName <> "ERROR" Then

                If MyHub <> "" And Not MyHub.ToUpper.StartsWith("HO") And Not MyHub.ToUpper.StartsWith("FA") Then
                    dt.DefaultView.RowFilter = "Value = '" & MyHub & "' or Value = ''" 'khusus hub sendiri
                End If

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetHubListNoFilter() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.GetHubList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetStationHubList(ByVal PageName As String, ByVal MyStationHub As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetStationHubList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                If MyStationHub <> "" And Not MyStationHub.ToUpper.StartsWith("HO") Then
                    dt.DefaultView.RowFilter = "Value = '" & MyStationHub & "' or Value = ''" 'khusus hub sendiri
                End If

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetServiceTypeList(ByVal PageName As String, ByVal MyHub As String, Optional ByVal ShowSvcCode As Boolean = False) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = Nothing 'ServiceIna
            param(3) = ShowSvcCode

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.GetServiceTypeList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetAccountCategoryList() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetAccountCategoryList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetAccountSubCategoryList() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetAccountSubCategoryList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetECommerceList(ByVal PageName As String, ByVal MyHub As String) As DataTable

        Return GetECommerceList(PageName, MyHub, "", "", "")

    End Function

    Public Function GetECommerceList(ByVal PageName As String, ByVal MyHub As String, ByVal ServiceType As String) As DataTable

        Return GetECommerceList(PageName, MyHub, ServiceType, "", "")

    End Function

    Public Function GetECommerceList(ByVal PageName As String, ByVal MyHub As String, ByVal ServiceType As String, ByVal ServiceAddInfo As String, ByVal OtherCriteria As String, Optional ByVal ShowAccountCode As Boolean = False) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = ServiceType
            param(3) = ServiceAddInfo
            param(4) = OtherCriteria
            param(5) = ShowAccountCode

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.GetECommerceList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetECommerceListWithCategoryAndSubCategory(ByVal PageName As String, ByVal MyHub As String, ByRef dsData As DataSet) As DataTable

        Return GetECommerceListWithCategoryAndSubCategory("", "", "", "", "", False, dsData)

    End Function

    Public Function GetECommerceListWithCategoryAndSubCategory(ByVal PageName As String, ByVal MyHub As String, ByVal ServiceType As String, ByVal ServiceAddInfo As String, ByVal OtherCriteria As String, ByVal ShowAccCode As Boolean, ByRef dsData As DataSet) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = ServiceType
            param(3) = ServiceAddInfo
            param(4) = OtherCriteria
            param(5) = ShowAccCode

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.GetECommerceListWithCategoryAndSubCategory(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetSuppAddrCodeList() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetSuppAddrCodeList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

    Public Function GetFleetRentalVehicleTypeListByAccount(ByVal PageName As String, ByVal MyAccount As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = MyAccount

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetFleetRentalVehicleTypeListByAccount(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatableWithBarcode(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetExpeditionExpenseTypeList(ByVal PageName As String, ByVal MyHub As String, Optional ByVal AllowAllSelection As Boolean = False) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            dt = New DataTable
            dt.TableName = "DATA"
            dt.Columns.Add("Value")
            dt.Columns.Add("Display")

            If AllowAllSelection Then
                dt.Rows.Add(New String() {"", ""}) 'ada proses / laporan yang boleh semua pilihan sekaligus
            End If
            dt.Rows.Add(New String() {"CON", "SERAH EKSPEDISI"})
            dt.Rows.Add(New String() {"AWB", "HUB ANTAR ALAMAT"})
            dt.Rows.Add(New String() {"PUE", "JEMPUT EKSPEDISI"})
            dt.Rows.Add(New String() {"PUP", "JEMPUT REKANAN"})
            dt.Rows.Add(New String() {"FLT", "SEWA ARMADA"})

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetExpeditionList(ByVal PageName As String, ByVal MyHub As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetExpeditionList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetExpeditionAutoOrderSetting(ByVal KodeHub As String, ByVal KodeEkspedisi As String, ByVal PageTimeout As Integer, ByRef ErrorMessage As String) As DataTable

        Dim dtLayanan As DataTable = Nothing
        Dim dtJenisBarang As DataTable = Nothing

        Return GetExpeditionAutoOrderSetting("HEADER", KodeHub, KodeEkspedisi, PageTimeout, dtLayanan, dtJenisBarang, ErrorMessage)

    End Function

    Public Function GetExpeditionAutoOrderSetting(ByVal KodeHub As String, ByVal KodeEkspedisi As String, ByVal PageTimeout As Integer, ByRef dtLayanan As DataTable, ByRef dtJenisBarang As DataTable, ByRef ErrorMessage As String) As DataTable

        Return GetExpeditionAutoOrderSetting("DETAIL", KodeHub, KodeEkspedisi, PageTimeout, dtLayanan, dtJenisBarang, ErrorMessage)

    End Function

    Private Function GetExpeditionAutoOrderSetting(ByVal ReturnType As String, ByVal KodeHub As String, ByVal KodeEkspedisi As String, ByVal PageTimeout As Integer, ByRef dtLayanan As DataTable, ByRef dtJenisBarang As DataTable, ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing
        Dim dtDetail As DataTable = Nothing

        Dim ds As DataSet = Nothing

        Try

            Dim param(3) As Object
            param(0) = UserWS
            param(1) = PassWS
            param(2) = KodeHub 'Hub Session("CurrentWebCode")
            param(3) = KodeEkspedisi 'Kode Ekspedisi

            Dim serv As New CoreService.Service
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            serv.Timeout = PageTimeout '3600000 'In miliseconds, 1 Jam

            Dim respon() As Object = Nothing

            Dim i As Integer = 1
            Dim success As Boolean = False
            While i <= maxTryWS And success = False

                Try

                    respon = serv.GetExpeditionAutoOrderSetting(AppName, AppVersion, param, ds)
                    success = True

                Catch ex As Exception

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                    i = i + 1

                End Try

            End While

            If success Then

                If respon(0) = "0" Then

                    If Not IsNothing(ds) Then
                        If ds.Tables.Count > 0 Then

                            Dim successHeader As Boolean = False
                            Dim successDetail As Boolean = False

                            For j As Integer = 0 To ds.Tables.Count - 1
                                If ds.Tables(j).TableName.ToUpper = "HEADER" And ReturnType = "HEADER" Then
                                    dt = ds.Tables(j)
                                    successHeader = True
                                    Exit For
                                ElseIf ds.Tables(j).TableName.ToUpper = "DETAIL" And ReturnType = "DETAIL" Then
                                    dtDetail = ds.Tables(j)
                                    successDetail = True
                                    Exit For
                                End If
                            Next

                            If ReturnType = "HEADER" Then

                                If successHeader Then
                                    If dt.Rows.Count > 0 Then
                                        If dt.Rows(0).Item("Expedition") <> "" Then
                                            Dim dr As DataRow = dt.NewRow
                                            dr("Expedition") = ""
                                            dr("ExpeditionName") = ""
                                            dt.Rows.InsertAt(dr, 0)
                                        End If
                                    Else
                                        ErrorMessage = "Tidak ada Ekspedisi!"
                                        Return Nothing
                                    End If
                                Else
                                    ErrorMessage = "Tidak ada table ""Header"" di DataSet!"
                                    Return Nothing
                                End If

                            ElseIf ReturnType = "DETAIL" Then

                                If successDetail Then
                                    If dtDetail.Rows.Count > 0 Then

                                        If dtDetail.Rows(0).Item("Value").ToString.ToUpper <> "NOTHING" Then

                                            dtDetail.DefaultView.RowFilter = "`Category` = 'Layanan'"
                                            dtLayanan = dtDetail.DefaultView.ToTable

                                            dtDetail.DefaultView.RowFilter = "`Category` = 'Jenis Barang'"
                                            dtJenisBarang = dtDetail.DefaultView.ToTable

                                            If dtLayanan.Rows.Count > 0 And dtJenisBarang.Rows.Count > 0 Then

                                                If dtLayanan.Rows(0).Item("Value") <> "" Then

                                                    Dim dr1 As DataRow = dtLayanan.NewRow
                                                    For i1 As Integer = 0 To dtLayanan.Columns.Count - 1
                                                        If dtLayanan.Columns(i1).DataType.Equals(System.Type.GetType("System.String")) Then
                                                            dr1(i1) = ""
                                                        End If
                                                    Next
                                                    dtLayanan.Rows.InsertAt(dr1, 0)

                                                End If

                                                If dtJenisBarang.Rows(0).Item("Value") <> "" Then

                                                    Dim dr2 As DataRow = dtJenisBarang.NewRow
                                                    For i2 As Integer = 0 To dtJenisBarang.Columns.Count - 1
                                                        If dtJenisBarang.Columns(i2).DataType.Equals(System.Type.GetType("System.String")) Then
                                                            dtJenisBarang.Rows.InsertAt(dr2, 0)
                                                        End If
                                                        dr2(i2) = ""
                                                    Next

                                                End If

                                            Else

                                                ErrorMessage = "Table Detail, Category"
                                                If dtLayanan.Rows.Count < 1 Then
                                                    ErrorMessage &= " ""Layanan"""
                                                End If
                                                If dtJenisBarang.Rows.Count < 1 Then
                                                    ErrorMessage &= " ""Jenis Barang"""
                                                End If
                                                ErrorMessage &= " tidak ada data"

                                                Return Nothing

                                            End If

                                        Else
                                            ErrorMessage = "Value dari Table ""Detail"" NOTHING!"
                                            Return Nothing
                                        End If


                                        If dt.Rows(0).Item("Expedition") <> "" Then
                                            Dim dr As DataRow = dt.NewRow
                                            dr("Expedition") = ""
                                            dr("ExpeditionName") = ""
                                            dt.Rows.InsertAt(dr, 0)
                                        End If
                                    Else
                                        ErrorMessage = "Tidak ada data Detail Ekspedisi!"
                                        Return Nothing
                                    End If
                                Else
                                    ErrorMessage = "Tidak ada table ""Detail"" di DataSet!"
                                    Return Nothing
                                End If

                            End If

                        Else
                            ErrorMessage = "Tidak ada table di DataSet!"
                            Return Nothing
                        End If
                    Else
                        ErrorMessage = respon(1)
                        Return Nothing
                    End If

                Else
                    ErrorMessage = respon(1)
                    Return Nothing
                End If

            Else

                ErrorMessage = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                WriteTracelogTxt(Pesan)

                Return Nothing

            End If

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            dt = Nothing
            dtDetail = Nothing
            ds = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function GetIndogrosirStoreList(ByVal pType As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_indogrosirstore_list`(@pType);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@pType", pType)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetIndomaretDCList(ByVal FirstRowEmptyValue As Boolean, ByVal IncludeIgr As Boolean) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = "INDOPAKET"
            param(3) = IIf(IncludeIgr, "1", "0")

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetIndomaretDCList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("Code") <> "" Then

                        If FirstRowEmptyValue Then
                            Dim dr As DataRow = dt.NewRow
                            dr("Code") = ""
                            dr("Name") = ""

                            dt.Rows.InsertAt(dr, 0)
                        End If

                    End If
                End If

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetIndomaretDCDepoList(ByVal FirstRowEmptyValue As Boolean) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = "INDOPAKET"

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetIndomaretDCDepoList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("Value") <> "" Then

                        If FirstRowEmptyValue Then
                            Dim dr As DataRow = dt.NewRow
                            dr("Value") = ""
                            dr("Display") = ""

                            dt.Rows.InsertAt(dr, 0)
                        End If

                    End If
                End If

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetRackInStoreList(ByVal DCCode As String, ByVal StoreList As String, ByVal TrackNumList As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(6)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = StoreList.Replace(",", "|")
            param(3) = TrackNumList.Replace(",", "|")

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetRackInStoreList(ClsWebVer.AppName, ClsWebVer.AppVersion, DCCode, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetPackageConditionList(ByVal PageName As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetPackageConditionList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetECommerceListByFinType(ByVal AccountFinType As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = AccountFinType

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetECommerceListByFinType(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetThirdPartyList() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetThirdPartyList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DCPickingManual(ByVal DCCode As String, ByVal StoreList As String, ByVal TrackNumList As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim serv2 As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(6)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = StoreList.Replace(",", "|")
            param(6) = TrackNumList.Replace(",", "|")

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv2.DCPickingManual(ClsWebVer.AppName, ClsWebVer.AppVersion, DCCode, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function


    Public Function IndomaretStoreClosedList(ByVal StoreList As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = StoreList

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.IndomaretStoreClosedList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IndomaretStoreClosedOpen(ByVal StoreList As String, ByRef PesanError As String) As Boolean

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = StoreList

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.IndomaretStoreClosedOpen(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                PesanError = respon(1)
                Return True

            Else

                PesanError = respon(1)
                Return False

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    'Public Function PrintSuratJalan(ByVal NoSuratJalan As String, ByVal TipePenerima As String, ByVal TipeCetak As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByRef PesanError As String) As Boolean

    '    Dim CreatedFileName As String = ""
    '    Return PrintSuratJalan(NoSuratJalan, TipePenerima, TipeCetak, CurrentWebCode, CurrentWebName, CreatedFileName, PesanError)

    'End Function

    'Public Function PrintSuratJalan(ByVal NoSuratJalan As String, ByVal TipePenerima As String, ByVal TipeCetak As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByRef CreatedFileName As String, ByRef PesanError As String) As Boolean

    '    If CurrentWebCode <> "" Then

    '        Dim NamaRpt As String = "SuratJalan.rpt"
    '        Dim dt As New DataTable("SuratJalan")

    '        Try

    '            Dim IsDelivery As String = "0"
    '            If TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                IsDelivery = "1"
    '            End If

    '            Dim serv As New CoreService.Service
    '            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

    '            Dim respon() As Object = Nothing

    '            Dim param(5) As Object
    '            param(0) = UserWS
    '            param(1) = PassWS
    '            param(2) = "HUB"
    '            param(3) = CurrentWebCode
    '            param(4) = NoSuratJalan
    '            param(5) = IsDelivery

    '            Dim i As Integer = 1
    '            Dim success As Boolean = False
    '            While i <= maxTryWS And success = False

    '                Try

    '                    respon = serv.PrintSuratJalan(AppName, AppVersion, param)
    '                    success = True

    '                Catch ex As Exception

    '                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                    Dim Pesan As String = ""
    '                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
    '                    WriteTracelogTxt(Pesan)

    '                    i = i + 1

    '                End Try

    '            End While

    '            If success Then

    '                If respon(0) = "0" Then

    '                    Dim header() As String = Nothing
    '                    header = respon(1).ToString.Split("|")
    '                    Dim isCons As String = "0"
    '                    Try
    '                        isCons = header(14)
    '                    Catch ex As Exception
    '                        isCons = "0"
    '                    End Try

    '                    If isCons = "1" Then
    '                        If PrintSuratJalan_Rangkuman(respon, NoSuratJalan, TipePenerima, TipeCetak, CurrentWebCode, CurrentWebName, CreatedFileName, PesanError) Then
    '                        Else
    '                            Return False
    '                        End If
    '                    End If

    '                    If TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                        dt = ConvertStringToDatatableWithBarcode(respon(2))
    '                    Else
    '                        dt = ConvertStringToDatatableForSJ(respon(2))
    '                    End If

    '                    'Kalau Print Surat Jalan Roti, Ambil Alih disini
    '                    If TipePenerima = "DCR" Or TipePenerima = "DCI" _
    '                    Or TipePenerima = "IGR" Or TipePenerima = "H2H" _
    '                    Or TipePenerima = "P2H" Then
    '                        Try
    '                            If PrintSuratJalan_ByDst(respon, dt.DefaultView.ToTable(), "", NoSuratJalan, TipePenerima, TipeCetak, CurrentWebCode, CurrentWebName, CreatedFileName, PesanError) Then
    '                                Return True
    '                            Else
    '                                Return False
    '                            End If
    '                        Catch ex As Exception
    '                            PesanError = ex.Message

    '                            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                            Dim Pesan As String = ""
    '                            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '                            WriteTracelogTxt(Pesan)

    '                            Return False
    '                        End Try
    '                    End If
    '                    'Print Surat Jalan Roti selesai, code dibawahnya tidak perlu dijalankan

    '                    Dim dv As New DataView(dt)
    '                    Dim dt_tujuan As New DataTable()
    '                    Dim columnNames(0) As String
    '                    If TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                        columnNames(0) = "Zone"
    '                    Else
    '                        columnNames(0) = "Dst"
    '                    End If

    '                    dt_tujuan = dv.ToTable(True, columnNames)

    '                    Dim Index As Integer = 0
    '                    Dim TotalDst As Integer = dt_tujuan.Rows.Count - 1
    '                    Dim DstCode As String = ""

    '                    Dim SB As New StringBuilder

    '                    While Index <= TotalDst

    '                        CreatedFileName = ""

    '                        If TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                            DstCode = dt_tujuan.Rows(Index).Item("Zone")
    '                            dt.DefaultView.RowFilter = "Zone = '" & DstCode & "'"
    '                        Else
    '                            DstCode = dt_tujuan.Rows(Index).Item("Dst")
    '                            dt.DefaultView.RowFilter = "Dst = '" & DstCode & "'"
    '                        End If

    '                        If PrintSuratJalan_ByDst(respon, dt.DefaultView.ToTable(), DstCode, NoSuratJalan, TipePenerima, TipeCetak, CurrentWebCode, CurrentWebName, CreatedFileName, PesanError) Then
    '                            If Index > 0 Then
    '                                SB.Append("|")
    '                            End If
    '                            SB.Append(CreatedFileName)
    '                        Else
    '                            Return False
    '                        End If

    '                        dt.DefaultView.RowFilter = Nothing

    '                        Index = Index + 1

    '                    End While

    '                    CreatedFileName = SB.ToString

    '                    Return True

    '                Else

    '                    PesanError = respon(1)

    '                    Return False

    '                End If

    '            Else

    '                PesanError = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

    '                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                Dim Pesan As String = ""
    '                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
    '                WriteTracelogTxt(Pesan)

    '                Return False

    '            End If

    '            Return False

    '        Catch ex As Exception

    '            PesanError = ex.Message

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '            WriteTracelogTxt(Pesan)

    '            Return False

    '        Finally

    '            dt = Nothing
    '            GC.Collect()

    '        End Try

    '    End If

    'End Function

    'Private Function PrintSuratJalan_Rangkuman(ByVal ori_respon() As Object, ByVal NoSuratJalan As String, ByVal TipePenerima As String, ByVal TipeCetak As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByRef CreatedFileName As String, ByRef PesanError As String) As Boolean

    '    If CurrentWebCode <> "" Then

    '        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '        Dim NamaRpt As String = "SuratJalanRangkuman.rpt"
    '        Dim dt As New DataTable("SuratJalan")

    '        Try

    '            Dim MyRespon(3) As Object
    '            MyRespon = ori_respon

    '            If MyRespon(0) = "0" Then

    '                Dim header() As String = Nothing
    '                header = MyRespon(1).ToString.Split("|")

    '                Dim respCompanyName As String = header(0)
    '                Dim respCompanyAddress As String = header(1)
    '                Dim respCompanyPhone As String = header(2)
    '                Dim respCompanyFax As String = header(3)
    '                Dim SjNum As String = header(4)
    '                Dim SjNumH As String = header(5)
    '                Dim DriverID As String = header(6)
    '                Dim respDriverName As String = header(7)
    '                Dim Vehicle As String = header(8)
    '                Dim Expedition As String = header(9)
    '                Dim Type As String = header(10)
    '                Dim Code As String = header(11)
    '                Dim Jumlah As String = header(12)
    '                Dim CreatedDate As String = header(13)

    '                Dim isCons As String = "0"
    '                Try
    '                    isCons = header(14)
    '                Catch ex As Exception
    '                    isCons = "0"
    '                End Try

    '                Dim SenderID As String = header(15)
    '                Dim SenderName As String = header(16)

    '                Dim bSjNum As String = MyRespon(3).ToString()

    '                Dim CompanyName As String = respCompanyName
    '                Dim CompanyAddress As String = respCompanyAddress

    '                Dim CompanyPhone As String = ""
    '                If respCompanyPhone.Trim <> "" Then
    '                    CompanyPhone = "Phone : " & respCompanyPhone
    '                End If

    '                Dim CompanyFax As String = ""
    '                If respCompanyFax.Trim <> "" Then
    '                    CompanyFax = "Fax : " & respCompanyFax
    '                End If

    '                Dim strReprint As String = ""
    '                If TipeCetak.ToUpper = "REPRINT" Then
    '                    strReprint = "REPRINT "
    '                End If

    '                Dim DocumentName As String = strReprint & "RANGKUMAN SURAT JALAN " & Type
    '                Dim DocumentNo As String = "NO. " & SjNum & " / " & Code & "-" & CurrentWebName
    '                Dim PackageType As String = "Nomor Kons"
    '                Dim Pengirim As String = SenderName
    '                Dim PengirimID As String = SenderID
    '                Dim Penerima As String = "...................."
    '                Dim TotalPackage As String = Jumlah
    '                Dim BarcodeDocumentNo As String = bSjNum
    '                Dim TanggalDocument As String = CreatedDate
    '                Dim ExpeditionName As String = Expedition
    '                Dim VehicleNo As String = Vehicle
    '                Dim Origin As String = Code & "-" & CurrentWebName
    '                Dim DriverName As String = respDriverName

    '                dt = ConvertStringToDatatableForSJ(MyRespon(2))

    '                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

    '                rpt.Load(FileRptPath)
    '                rpt.SummaryInfo.ReportTitle = "Surat Jalan"
    '                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
    '                rpt.SetDataSource(dt)

    '                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
    '                'Dim crParameterFieldLocation As ParameterFieldDefinition
    '                Dim crParameterValues As ParameterValues
    '                Dim crParameterDiscreteValue As ParameterDiscreteValue

    '                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyAddress")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyAddress")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyAddress
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyPhone")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyPhone")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyPhone
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyFax")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyFax")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyFax
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DocumentName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DocumentNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PackageType")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("PackageType")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = PackageType
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Pengirim")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Pengirim")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Pengirim
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PengirimID")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("PengirimID")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = PengirimID
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Penerima")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Penerima")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Penerima
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TotalPackage")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("TotalPackage")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = TotalPackage
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = BarcodeDocumentNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TanggalDocument")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalDocument")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = TanggalDocument
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ExpeditionName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("ExpeditionName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = ExpeditionName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("VehicleNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("VehicleNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = VehicleNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Origin")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Origin")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Origin
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DriverName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverID")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverID")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DriverID
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    Dim ShowLogo As String = ReadWebConfig("RPTShowLogo", False)
    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ShowLogo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("ShowLogo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = ShowLogo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                End Using

    '                Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
    '                Dim FullPathReportFile As String = PathReportFile & "\SuratJalan_" & CurrentWebCode & "_NoSJ_" & NoSuratJalan & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmss") & "_Rangkuman.pdf"
    '                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

    '                Try

    '                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
    '                    Dim PrintToPDF As String = "1"
    '                    If PrintToPDF = "1" Then
    '                        If Not Directory.Exists(PathReportFile) Then
    '                            Directory.CreateDirectory(PathReportFile)
    '                        End If

    '                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)

    '                        CreatedFileName = FullPathReportFile

    '                    End If

    '                Catch ex As Exception

    '                    PesanError = ex.Message

    '                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                    Dim Pesan As String = ""
    '                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
    '                    WriteTracelogTxt(Pesan)

    '                End Try

    '                Try

    '                    Dim PrintToPrinter As String = ReadWebConfig("PrintToPrinter", False)

    '                    If PrintToPrinter = "1" Then
    '                        Dim HubPrinterName As String = ReadWebConfig("PrinterName", False)
    '                        rpt.PrintOptions.PrinterName = HubPrinterName
    '                        rpt.PrintToPrinter(1, True, 0, 0)
    '                    End If

    '                Catch

    '                    Using prd As System.Drawing.Printing.PrintDocument = New System.Drawing.Printing.PrintDocument

    '                        'Dim prd As New System.Drawing.Printing.PrintDocument
    '                        rpt.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName
    '                        rpt.PrintToPrinter(1, True, 0, 0)

    '                    End Using

    '                End Try

    '                'Try
    '                '    crParameterFieldDefinitions.Dispose()
    '                'Catch ex As Exception
    '                'End Try

    '                'Try
    '                '    crParameterFieldLocation.Dispose()
    '                'Catch ex As Exception
    '                'End Try

    '                Return True

    '            Else
    '                PesanError = MyRespon(1)

    '                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                Dim Pesan As String = ""
    '                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & PesanError
    '                WriteTracelogTxt(Pesan)

    '                Return False
    '            End If

    '            Return False

    '        Catch ex As Exception

    '            PesanError = ex.Message

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '            WriteTracelogTxt(Pesan)

    '            Return False

    '        Finally
    '            Try
    '                rpt.Close()
    '            Catch ex As Exception
    '            End Try
    '            Try
    '                rpt.Dispose()
    '            Catch ex As Exception
    '            End Try

    '            dt = Nothing
    '            GC.Collect()

    '        End Try

    '    End If

    'End Function

    'Private Function PrintSuratJalan_ByDst(ByVal ori_respon() As Object, ByVal dt_respon2 As DataTable, ByVal DstCode As String, ByVal NoSuratJalan As String, ByVal TipePenerima As String, ByVal TipeCetak As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByRef CreatedFileName As String, ByRef PesanError As String) As Boolean

    '    If CurrentWebCode <> "" Then

    '        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    '        Dim NamaRpt As String = "SuratJalanCons.rpt"
    '        Dim dt As New DataTable("SuratJalan")

    '        Try

    '            Dim MyRespon(3) As Object
    '            MyRespon = ori_respon

    '            If MyRespon(0) = "0" Then

    '                Dim header() As String = Nothing
    '                header = MyRespon(1).ToString.Split("|")

    '                Dim respCompanyName As String = header(0)
    '                Dim respCompanyAddress As String = header(1)
    '                Dim respCompanyPhone As String = header(2)
    '                Dim respCompanyFax As String = header(3)
    '                Dim SjNum As String = header(4)
    '                Dim SjNumH As String = header(5)
    '                Dim DriverID As String = header(6)
    '                Dim respDriverName As String = header(7)
    '                Dim Vehicle As String = header(8)
    '                Dim Expedition As String = header(9)
    '                Dim Type As String = header(10)
    '                Dim Code As String = header(11)
    '                Dim Jumlah As String = header(12)
    '                Dim CreatedDate As String = header(13)

    '                Dim isCons As String = "0"
    '                Try
    '                    isCons = header(14)
    '                Catch
    '                    isCons = "0"
    '                End Try

    '                Dim SenderID As String = header(15)
    '                Dim SenderName As String = header(16)

    '                Dim RefNo As String = ""
    '                Try
    '                    RefNo = header(17)
    '                Catch
    '                    RefNo = ""
    '                End Try

    '                'header(18) = Urut By Zona atau tidak

    '                Dim DriverNotes As String = ""
    '                Try
    '                    DriverNotes = header(19)
    '                Catch
    '                End Try

    '                Dim bSjNum As String = MyRespon(3).ToString()

    '                Dim CompanyName As String = respCompanyName
    '                Dim CompanyAddress As String = respCompanyAddress

    '                Dim CompanyPhone As String = ""
    '                If respCompanyPhone.Trim <> "" Then
    '                    CompanyPhone = "Phone : " & respCompanyPhone
    '                End If

    '                Dim CompanyFax As String = ""
    '                If respCompanyFax.Trim <> "" Then
    '                    CompanyFax = "Fax : " & respCompanyFax
    '                End If

    '                Dim strReprint As String = ""
    '                If TipeCetak.ToUpper = "REPRINT" Then
    '                    strReprint = "REPRINT "
    '                End If

    '                Dim DocumentName As String = strReprint & "DETAIL SURAT JALAN " & Type
    '                If TipePenerima = "DCR" Or TipePenerima = "DCI" _
    '                Or TipePenerima = "IGR" Or TipePenerima = "H2H" _
    '                Or TipePenerima = "P2H" Then
    '                    DocumentName = strReprint & "SURAT JALAN PAKET"
    '                ElseIf TipePenerima = "H2C" Then
    '                    DocumentName = strReprint & "SURAT JALAN DELIVERY"
    '                ElseIf TipePenerima = "HIP" Then
    '                    DocumentName = strReprint & "DAFTAR JEMPUT"
    '                ElseIf TipePenerima = "JPT" Then
    '                    DocumentName = strReprint & "DAFTAR JEMPUT PAKET DI TOKO"
    '                End If

    '                Dim DocumentNo As String = "NO. " & SjNum & " / " & Code & "-" & CurrentWebName
    '                If TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                    DocumentNo &= " / Zona " & DstCode
    '                End If

    '                Dim PackageType As String = "Nomor Kons"
    '                Dim Pengirim As String = SenderName
    '                Dim PengirimID As String = SenderID
    '                Dim Penerima As String = "...................."
    '                Dim TotalPackage As String = Jumlah
    '                Dim BarcodeDocumentNo As String = bSjNum
    '                Dim TanggalDocument As String = CreatedDate
    '                Dim ExpeditionName As String = Expedition
    '                Dim VehicleNo As String = Vehicle
    '                Dim Origin As String = Code & "-" & CurrentWebName
    '                Dim DriverName As String = respDriverName

    '                If TipePenerima = "DCR" Or TipePenerima = "DCI" Or TipePenerima = "IGR" Or TipePenerima = "JPT" Then
    '                    NamaRpt = "SuratJalanRoti.rpt"
    '                ElseIf TipePenerima = "H2H" Or TipePenerima = "P2H" Then
    '                    NamaRpt = "SuratJalanDGA.rpt"
    '                ElseIf TipePenerima = "H2C" Or TipePenerima = "HIP" Then
    '                    NamaRpt = "SuratJalanDelivery.rpt"
    '                Else
    '                    If isCons = "0" Then
    '                        PackageType = "Nomor Resi"
    '                        NamaRpt = "SuratJalan.rpt"
    '                    End If
    '                End If

    '                dt = dt_respon2

    '                Try
    '                    Dim dtA As New DataTable("SuratJalanA")
    '                    If MyRespon(4).ToString <> "" Then
    '                        dtA = ConvertStringToDatatableForSJ(MyRespon(4).ToString)
    '                        dt.Merge(dtA)
    '                    End If
    '                Catch
    '                    'Hanya untuk jaga2 kalau belum ada respon(4)
    '                End Try

    '                If TipePenerima = "RPX" Then

    '                    Dim Index As Integer = 0
    '                    Dim MaxIndex As Integer = dt.Rows.Count - 1
    '                    While Index <= MaxIndex

    '                        dt.Rows(Index).Item("Dst") = "RPX"
    '                        dt.Rows(Index).Item("DstName") = "RPX"

    '                        Index = Index + 1

    '                    End While

    '                End If

    '                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

    '                rpt.Load(FileRptPath)
    '                rpt.SummaryInfo.ReportTitle = "Surat Jalan"
    '                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
    '                rpt.SetDataSource(dt)

    '                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
    '                'Dim crParameterFieldLocation As ParameterFieldDefinition
    '                Dim crParameterValues As ParameterValues
    '                Dim crParameterDiscreteValue As ParameterDiscreteValue

    '                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyAddress")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyAddress")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyAddress
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyPhone")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyPhone")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyPhone
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyFax")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyFax")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = CompanyFax
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DocumentName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DocumentNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PackageType")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("PackageType")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = PackageType
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Pengirim")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Pengirim")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Pengirim
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PengirimID")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("PengirimID")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = PengirimID
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Penerima")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Penerima")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Penerima
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TotalPackage")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("TotalPackage")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = TotalPackage
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = BarcodeDocumentNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TanggalDocument")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalDocument")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = TanggalDocument
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ExpeditionName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("ExpeditionName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = ExpeditionName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("VehicleNo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("VehicleNo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = VehicleNo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Origin")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("Origin")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = Origin
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverName")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverName")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DriverName
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverID")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverID")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = DriverID
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    Dim ShowLogo As String = ReadWebConfig("RPTShowLogo", False)
    '                    'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ShowLogo")
    '                        'crParameterFieldLocation = crParameterFieldDefinitions.Item("ShowLogo")
    '                        crParameterValues = crParameterFieldLocation.CurrentValues
    '                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                        crParameterDiscreteValue.Value = ShowLogo
    '                        crParameterValues.Add(crParameterDiscreteValue)
    '                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                    End Using

    '                    If TipePenerima = "DCR" Or TipePenerima = "DCI" _
    '                    Or TipePenerima = "IGR" Or TipePenerima = "H2H" _
    '                    Or TipePenerima = "P2H" Or TipePenerima = "JPT" Then
    '                        'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                        Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("RefNo")
    '                            'crParameterFieldLocation = crParameterFieldDefinitions.Item("RefNo")
    '                            crParameterValues = crParameterFieldLocation.CurrentValues
    '                            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                            crParameterDiscreteValue.Value = RefNo
    '                            crParameterValues.Add(crParameterDiscreteValue)
    '                            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                        End Using


    '                        If TipePenerima = "DCR" Or TipePenerima = "DCI" Or TipePenerima = "IGR" _
    '                        Or TipePenerima = "JPT" Then
    '                            'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
    '                            Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverNotes")
    '                                'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverNotes")
    '                                crParameterValues = crParameterFieldLocation.CurrentValues
    '                                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
    '                                crParameterDiscreteValue.Value = DriverNotes
    '                                crParameterValues.Add(crParameterDiscreteValue)
    '                                crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
    '                            End Using
    '                        End If

    '                    End If

    '                End Using

    '                Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
    '                Dim NamaSuratJalan As String = "SuratJalan"
    '                If TipePenerima = "DCR" Then
    '                    NamaSuratJalan = "SuratJalanRoti"
    '                ElseIf TipePenerima = "JPT" Then
    '                    NamaSuratJalan = "JemputPaketDiToko"
    '                Else
    '                    NamaSuratJalan = "SuratJalan" & TipePenerima
    '                End If
    '                If DstCode <> "" Then
    '                    DstCode = "_Dst_" & DstCode
    '                End If
    '                Dim FullPathReportFile As String = PathReportFile & "\" & NamaSuratJalan & "_" & CurrentWebCode & "_NoSJ_" & NoSuratJalan & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmss") & DstCode & ".pdf"
    '                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

    '                Try
    '                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
    '                    Dim PrintToPDF As String = "1"
    '                    If PrintToPDF = "1" Then
    '                        If Not Directory.Exists(PathReportFile) Then
    '                            Directory.CreateDirectory(PathReportFile)
    '                        End If

    '                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
    '                        CreatedFileName = FullPathReportFile
    '                    End If
    '                Catch ex As Exception

    '                    PesanError = ex.Message

    '                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                    Dim Pesan As String = ""
    '                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
    '                    WriteTracelogTxt(Pesan)

    '                End Try

    '                Try

    '                    Dim PrintToPrinter As String = ReadWebConfig("PrintToPrinter", False)

    '                    If PrintToPrinter = "1" Then
    '                        Dim HubPrinterName As String = ReadWebConfig("PrinterName", False)
    '                        rpt.PrintOptions.PrinterName = HubPrinterName
    '                        rpt.PrintToPrinter(1, True, 0, 0)
    '                    End If

    '                Catch ex As Exception

    '                    Using prd As System.Drawing.Printing.PrintDocument = New System.Drawing.Printing.PrintDocument

    '                        'Dim prd As New System.Drawing.Printing.PrintDocument
    '                        rpt.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName
    '                        rpt.PrintToPrinter(1, True, 0, 0)

    '                    End Using

    '                End Try

    '                'Try
    '                '    crParameterFieldDefinitions.Dispose()
    '                'Catch ex As Exception
    '                'End Try

    '                'Try
    '                '    crParameterFieldLocation.Dispose()
    '                'Catch ex As Exception
    '                'End Try

    '                Return True
    '            Else
    '                PesanError = MyRespon(1)

    '                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '                Dim Pesan As String = ""
    '                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & PesanError
    '                WriteTracelogTxt(Pesan)

    '                Return False
    '            End If

    '            Return False

    '        Catch ex As Exception
    '            PesanError = ex.Message

    '            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '            Dim Pesan As String = ""
    '            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.StackTrace
    '            WriteTracelogTxt(Pesan)

    '            Return False
    '        Finally

    '            Try
    '                rpt.Close()
    '            Catch ex As Exception
    '            End Try
    '            Try
    '                rpt.Dispose()
    '            Catch ex As Exception
    '            End Try

    '            dt = Nothing
    '            GC.Collect()

    '        End Try
    '    End If

    'End Function

    Public Function PrintUMDPengajuan(ByVal Header() As String, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenUMDPengajuan.rpt"

                Dim DocumentName As String = Header(0)
                Dim NoUMD As String = Header(1)
                Dim Lokasi As String = Header(2)
                Dim JenisBiaya As String = Header(3)
                Dim TglPengajuan As String = Header(4)
                Dim TglDibutuhkan As String = Header(5)
                Dim TglCetak As String = Header(6)
                Dim Nominal As String = Header(7)
                Dim Keterangan As String = Header(8)
                Dim Identitas1TTD1 As String = Header(9)
                Dim Identitas2TTD1 As String = Header(10)
                Dim Identitas1TTD2 As String = Header(11)
                Dim Identitas2TTD2 As String = Header(12)
                Dim UserName As String = Header(13)
                Dim UserID As String = Header(14)
                Dim PemohonName As String = Header(15)
                Dim PemohonID As String = Header(16)
                Dim Rekening As String = Header(17)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "DokumenPengajuanUMD"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                'rpt.SetDataSource(Detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocumentName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoUMD")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoUMD
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Lokasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Lokasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("JenisBiaya")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = JenisBiaya
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPengajuan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPengajuan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglDibutuhkan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglDibutuhkan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Nominal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Nominal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Keterangan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Keterangan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Rekening")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Rekening
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileUMD", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                If DocumentName.Contains("REPRINT") Then
                    FullPathReportFile &= "Reprint"
                End If
                FullPathReportFile &= "Pengajuan_" & CurrentNIK & "_" & NoUMD & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                SqlQuery = "Insert Ignore Into PrintDocument ("
                SqlQuery &= " IDNumber, Type, SeqNo, `Status`"
                SqlQuery &= " , PrintCompany, PrintBy, PrintID"
                SqlQuery &= " , UpdTime, UpdUser"
                SqlQuery &= " ) Select u.NoUMD, @Type, IfNull(Max(p.SeqNo),0) + 1, 1"
                SqlQuery &= " , @PrintCompany, @PrintBy, @PrintID"
                SqlQuery &= " , now(), @User"
                SqlQuery &= " From `MstUMD` u"
                SqlQuery &= " Left Join PrintDocument p on (u.NoUMD = p.IdNumber and p.Type = @Type)"
                SqlQuery &= " Where u.NoUMD = @NoUMD"
                SqlQuery &= " Group By u.NoUMD"

                SqlParam.Add("@PrintCompany", "IPP")
                SqlParam.Add("@PrintBy", "RPT")
                SqlParam.Add("@PrintID", CurrentNIK)
                SqlParam.Add("@User", CurrentNIK)
                SqlParam.Add("@NoUMD", NoUMD)
                SqlParam.Add("@Type", "Pengajuan")

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    Public Function PrintUMDCairBatal(ByVal Header() As String, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenUMDCairBatal.rpt"

                Dim DocumentName As String = Header(0)
                Dim Type As String = Header(1)
                Dim NoUMD As String = Header(2)
                Dim Lokasi As String = Header(3)
                Dim JenisBiaya As String = Header(4)
                Dim TglPengajuan As String = Header(5)
                Dim RedaksiTgl As String = Header(6)
                Dim TglCairBatal As String = Header(7)
                Dim TglCetak As String = Header(8)
                Dim AddUser As String = Header(9)
                Dim Nominal As String = Header(10)
                Dim Keterangan As String = Header(11)
                Dim Redaksi1 As String = Header(12)
                Dim NominalAlasan As String = Header(13)
                Dim KeteranganBatal As String = Header(14)
                Dim Identitas1TTD1 As String = Header(15)
                Dim Identitas2TTD1 As String = Header(16)
                Dim Identitas1TTD2 As String = Header(17)
                Dim Identitas2TTD2 As String = Header(18)
                Dim UserName As String = Header(19)
                Dim UserID As String = Header(20)
                Dim PemohonName As String = Header(21)
                Dim PemohonID As String = Header(22)
                Dim Rekening As String = Header(23)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "Dokumen" & Type & "UMD"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                'rpt.SetDataSource(Detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocumentName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoUMD")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoUMD
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Lokasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Lokasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("JenisBiaya")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = JenisBiaya
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPengajuan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPengajuan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("RedaksiTgl")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = RedaksiTgl
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCairBatal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCairBatal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("AddUser")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = AddUser
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Nominal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Nominal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Keterangan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Keterangan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Redaksi1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Redaksi1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NominalAlasan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NominalAlasan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("KeteranganBatal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = KeteranganBatal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Rekening")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Rekening
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileUMD", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                If DocumentName.Contains("REPRINT") Then
                    FullPathReportFile &= "Reprint"
                End If
                FullPathReportFile &= Type & "_" & CurrentNIK & "_" & NoUMD & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                SqlQuery = "Insert Ignore Into PrintDocument ("
                SqlQuery &= " IDNumber, Type, SeqNo, `Status`"
                SqlQuery &= " , PrintCompany, PrintBy, PrintID"
                SqlQuery &= " , UpdTime, UpdUser"
                SqlQuery &= " ) Select u.NoUMD, @Type, IfNull(Max(p.SeqNo),0) + 1, 1"
                SqlQuery &= " , @PrintCompany, @PrintBy, @PrintID"
                SqlQuery &= " , now(), @User"
                SqlQuery &= " From `MstUMD` u"
                SqlQuery &= " Left Join PrintDocument p on (u.NoUMD = p.IdNumber and p.Type = @Type)"
                SqlQuery &= " Where u.NoUMD = @NoUMD"
                SqlQuery &= " Group By u.NoUMD"

                SqlParam.Add("@PrintCompany", "IPP")
                SqlParam.Add("@PrintBy", "RPT")
                SqlParam.Add("@PrintID", CurrentNIK)
                SqlParam.Add("@User", CurrentNIK)
                SqlParam.Add("@NoUMD", NoUMD)
                SqlParam.Add("@Type", Type)

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    Public Function PrintUMDPtj(ByVal Header() As String, ByVal Detail As DataTable, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenUMDPtj.rpt"

                Dim DocumentName As String = Header(0)
                Dim NoUMD As String = Header(1)
                Dim Lokasi As String = Header(2)
                Dim JenisBiaya As String = Header(3)
                Dim TglPengajuan As String = Header(4)
                Dim TglPtj As String = Header(5)
                Dim TglRealisasi As String = Header(6)
                Dim TglCetak As String = Header(7)
                Dim Nominal As String = Header(8)
                Dim NominalCair As String = Header(9)
                Dim NominalRealisasi As String = Header(10)
                Dim NominalSisa As String = Header(11)
                Dim Keterangan As String = Header(12)
                Dim Identitas1TTD1 As String = Header(13)
                Dim Identitas2TTD1 As String = Header(14)
                Dim Identitas1TTD2 As String = Header(15)
                Dim Identitas2TTD2 As String = Header(16)
                Dim UserName As String = Header(17)
                Dim UserID As String = Header(18)
                Dim PemohonName As String = Header(19)
                Dim PemohonID As String = Header(20)
                Dim Rekening As String = Header(21)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "DokumenPtjUMD"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                rpt.SetDataSource(Detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocumentName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoUMD")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoUMD
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Lokasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Lokasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("JenisBiaya")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = JenisBiaya
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPengajuan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPengajuan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPtj")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPtj
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglRealisasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglRealisasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Nominal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Nominal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NominalCair")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NominalCair
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NominalRealisasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NominalRealisasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NominalSisa")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NominalSisa
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Keterangan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Keterangan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Rekening")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Rekening
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileUMD", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                If DocumentName.Contains("REPRINT") Then
                    FullPathReportFile &= "Reprint"
                End If
                FullPathReportFile &= "Ptj_" & CurrentNIK & "_" & NoUMD & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                SqlQuery = "Insert Ignore Into PrintDocument ("
                SqlQuery &= " IDNumber, Type, SeqNo, `Status`"
                SqlQuery &= " , PrintCompany, PrintBy, PrintID"
                SqlQuery &= " , UpdTime, UpdUser"
                SqlQuery &= " ) Select u.NoUMD, @Type, IfNull(Max(p.SeqNo),0) + 1, 1"
                SqlQuery &= " , @PrintCompany, @PrintBy, @PrintID"
                SqlQuery &= " , now(), @User"
                SqlQuery &= " From `MstUMD` u"
                SqlQuery &= " Left Join PrintDocument p on (u.NoUMD = p.IdNumber and p.Type = @Type)"
                SqlQuery &= " Where u.NoUMD = @NoUMD"
                SqlQuery &= " Group By u.NoUMD"

                SqlParam.Add("@PrintCompany", "IPP")
                SqlParam.Add("@PrintBy", "RPT")
                SqlParam.Add("@PrintID", CurrentNIK)
                SqlParam.Add("@User", CurrentNIK)
                SqlParam.Add("@NoUMD", NoUMD)
                SqlParam.Add("@Type", "Pertanggungjawaban")

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    Public Function PrintReimbPengajuan(ByVal Header() As String, ByVal detail As DataTable, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenReimbPengajuan.rpt"

                Dim DocumentName As String = Header(0)
                Dim NoRMB As String = Header(1)
                Dim Lokasi As String = Header(2)
                Dim JenisBiaya As String = Header(3)
                Dim TglPengajuan As String = Header(4)
                Dim TglCetak As String = Header(5)
                Dim Nominal As String = Header(6)
                Dim Keterangan As String = Header(7)
                Dim Identitas1TTD1 As String = Header(8)
                Dim Identitas2TTD1 As String = Header(9)
                Dim Identitas1TTD2 As String = Header(10)
                Dim Identitas2TTD2 As String = Header(11)
                Dim UserName As String = Header(12)
                Dim UserID As String = Header(13)
                Dim PemohonName As String = Header(14)
                Dim PemohonID As String = Header(15)
                Dim Rekening As String = Header(16)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "DokumenPengajuanReimbursement"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                rpt.SetDataSource(detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocumentName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoRMB")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoRMB
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Lokasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Lokasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("JenisBiaya")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = JenisBiaya
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPengajuan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPengajuan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Nominal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Nominal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Keterangan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Keterangan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Rekening")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Rekening
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileReimb", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                If DocumentName.Contains("REPRINT") Then
                    FullPathReportFile &= "Reprint"
                End If
                FullPathReportFile &= "Pengajuan_" & CurrentNIK & "_" & NoRMB & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                SqlQuery = "Insert Ignore Into PrintDocument ("
                SqlQuery &= " IDNumber, Type, SeqNo, `Status`"
                SqlQuery &= " , PrintCompany, PrintBy, PrintID"
                SqlQuery &= " , UpdTime, UpdUser"
                SqlQuery &= " ) Select u.NoRMB, @Type, IfNull(Max(p.SeqNo),0) + 1, 1"
                SqlQuery &= " , @PrintCompany, @PrintBy, @PrintID"
                SqlQuery &= " , now(), @User"
                SqlQuery &= " From `MstReimb` u"
                SqlQuery &= " Left Join PrintDocument p on (u.NoRMB = p.IdNumber and p.Type = @Type)"
                SqlQuery &= " Where u.NoRMB = @NoRMB"
                SqlQuery &= " Group By u.NoRMB"

                SqlParam.Add("@PrintCompany", "IPP")
                SqlParam.Add("@PrintBy", "RPT")
                SqlParam.Add("@PrintID", CurrentNIK)
                SqlParam.Add("@User", CurrentNIK)
                SqlParam.Add("@NoRMB", NoRMB)
                SqlParam.Add("@Type", "Pengajuan")

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    Public Function PrintReimbCairBatal(ByVal Header() As String, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenReimbCairBatal.rpt"

                Dim DocumentName As String = Header(0)
                Dim Type As String = Header(1)
                Dim NoRMB As String = Header(2)
                Dim Lokasi As String = Header(3)
                Dim JenisBiaya As String = Header(4)
                Dim TglPengajuan As String = Header(5)
                Dim RedaksiTgl As String = Header(6)
                Dim TglCairBatal As String = Header(7)
                Dim TglCetak As String = Header(8)
                Dim AddUser As String = Header(9)
                Dim Nominal As String = Header(10)
                Dim Keterangan As String = Header(11)
                Dim Redaksi1 As String = Header(12)
                Dim NominalAlasan As String = Header(13)
                Dim KeteranganBatal As String = Header(14)
                Dim Identitas1TTD1 As String = Header(15)
                Dim Identitas2TTD1 As String = Header(16)
                Dim Identitas1TTD2 As String = Header(17)
                Dim Identitas2TTD2 As String = Header(18)
                Dim UserName As String = Header(19)
                Dim UserID As String = Header(20)
                Dim PemohonName As String = Header(21)
                Dim PemohonID As String = Header(22)
                Dim Rekening As String = Header(23)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "Dokumen" & Type & "UMD"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                'rpt.SetDataSource(Detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocumentName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoRMB")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoRMB
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Lokasi")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Lokasi
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("JenisBiaya")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = JenisBiaya
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglPengajuan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglPengajuan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("RedaksiTgl")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = RedaksiTgl
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCairBatal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCairBatal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("AddUser")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = AddUser
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Nominal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Nominal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Keterangan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Keterangan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Redaksi1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Redaksi1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NominalAlasan")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NominalAlasan
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("KeteranganBatal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = KeteranganBatal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas1TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas1TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Identitas2TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Identitas2TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("UserID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = UserID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PemohonID")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = PemohonID
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Rekening")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = Rekening
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileReimb", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                If DocumentName.Contains("REPRINT") Then
                    FullPathReportFile &= "Reprint"
                End If
                FullPathReportFile &= Type & "_" & CurrentNIK & "_" & NoRMB & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                SqlQuery = "Insert Ignore Into PrintDocument ("
                SqlQuery &= " IDNumber, Type, SeqNo, `Status`"
                SqlQuery &= " , PrintCompany, PrintBy, PrintID"
                SqlQuery &= " , UpdTime, UpdUser"
                SqlQuery &= " ) Select u.NoRMB, @Type, IfNull(Max(p.SeqNo),0) + 1, 1"
                SqlQuery &= " , @PrintCompany, @PrintBy, @PrintID"
                SqlQuery &= " , now(), @User"
                SqlQuery &= " From `MstReimb` u"
                SqlQuery &= " Left Join PrintDocument p on (u.NoRMB = p.IdNumber and p.Type = @Type)"
                SqlQuery &= " Where u.NoRMB = @NoRMB"
                SqlQuery &= " Group By u.NoRMB"

                SqlParam.Add("@PrintCompany", "IPP")
                SqlParam.Add("@PrintBy", "RPT")
                SqlParam.Add("@PrintID", CurrentNIK)
                SqlParam.Add("@User", CurrentNIK)
                SqlParam.Add("@NoRMB", NoRMB)
                SqlParam.Add("@Type", Type)

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    Public Function PrintTGF(ByVal Header() As String, ByVal Detail As DataTable, ByVal CurrentNIK As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentNIK <> "" Then

            Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                Dim NamaRpt As String = "DokumenTGF.rpt"

                Dim DocName As String = Header(0)
                Dim NoTGF As String = Header(1)
                Dim FollowUpCode As String = Header(2)
                Dim FollowUp As String = Header(3)
                Dim TglCetak As String = Header(4)
                Dim TTD1Code As String = Header(5)
                Dim TTD2Code As String = Header(6)
                Dim TTD3Code As String = Header(7)
                Dim TTD4Code As String = Header(8)
                Dim TTD5Code As String = Header(9)
                Dim TTD1 As String = Header(10)
                Dim TTD2 As String = Header(11)
                Dim TTD3 As String = Header(12)
                Dim TTD4 As String = Header(13)
                Dim TTD5 As String = Header(14)
                Dim SellingTotal As String = Header(15)

                Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                rpt.Load(FileRptPath)
                rpt.SummaryInfo.ReportTitle = "DokumenTindakLanjutSGA"
                rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                rpt.SetDataSource(Detail)

                'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                'Dim crParameterFieldLocation As ParameterFieldDefinition
                Dim crParameterDiscreteValue As ParameterDiscreteValue
                Dim crParameterValues As ParameterValues

                Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocName")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = DocName
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NoTGF")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = NoTGF
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("FollowUp")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = FollowUp
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TglCetak")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TglCetak
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TTD1")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TTD1
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TTD2")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TTD2
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TTD3")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TTD3
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TTD4")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TTD4
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TTD5")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = TTD5
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                    Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("SellingTotal")
                        crParameterValues = crParameterFieldLocation.CurrentValues
                        crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                        crParameterDiscreteValue.Value = SellingTotal
                        crParameterValues.Add(crParameterDiscreteValue)
                        crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                    End Using

                End Using

                Dim PathReportFile As String = ReadWebConfig("PathReportFileTGF", False)
                Dim FullPathReportFile As String = PathReportFile & "\"
                FullPathReportFile &= "TGF_" & CurrentNIK & "_" & NoTGF & ".pdf"
                FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                Try
                    'Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                    Dim PrintToPDF As String = "1"
                    If PrintToPDF = "1" Then
                        If Not Directory.Exists(PathReportFile) Then
                            Directory.CreateDirectory(PathReportFile)
                        End If

                        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                        CreatedFilePath = FullPathReportFile
                    End If
                Catch ex As Exception

                    PesanError = ex.Message

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                End Try

                Dim SB As New StringBuilder

                For Each row As DataRow In Detail.Rows
                    If SB.Length > 0 Then
                        SB.Append("|")
                    End If
                    SB.Append(row.Item("TrackNum"))
                    SB.Append(",")
                    SB.Append(row.Item("TrxItemSeqno"))
                Next

                Dim Description As String = SB.ToString

                SqlQuery = "Insert Ignore Into TGAFollowUpPrint ("
                SqlQuery &= " ID, FollowUp, Description"
                SqlQuery &= " , DibuatOleh, Menyetujui"
                SqlQuery &= " , Mengetahui1, Mengetahui2, Mengetahui3"
                SqlQuery &= " , AddTime, AddUser"
                SqlQuery &= " ) Values ("
                SqlQuery &= " @NoTGF, @FollowUp, @Description"
                SqlQuery &= " , @DibuatOleh, @Menyetujui"
                SqlQuery &= " , @Mengetahui1, @Mengetahui2, @Mengetahui3"
                SqlQuery &= " , now(), @UserID"
                SqlQuery &= " )"

                SqlParam.Add("@NoTGF", NoTGF)
                SqlParam.Add("@FollowUp", FollowUp)
                SqlParam.Add("@Description", Description)
                SqlParam.Add("@DibuatOleh", TTD1Code)
                SqlParam.Add("@Menyetujui", TTD2Code)
                SqlParam.Add("@Mengetahui1", TTD3Code)
                SqlParam.Add("@Mengetahui2", TTD4Code)
                SqlParam.Add("@Mengetahui3", TTD5Code)
                SqlParam.Add("@UserID", CurrentNIK)

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

                Return True

            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                Return False

            Finally
                Try
                    rpt.Close()
                Catch ex As Exception
                End Try
                Try
                    rpt.Dispose()
                Catch ex As Exception
                End Try

                If Not MCon Is Nothing Then
                    If MCon.State <> ConnectionState.Closed Then
                        MCon.Close()
                    End If
                    MCon.Dispose()
                End If

                GC.Collect()
            End Try

        End If

        Return False

    End Function

    'Public Function GetLatLong(ByVal Address As String, ByVal PostalCode As String, ByRef ErrorMessage As String) As DataTable

    '    Dim dt As DataTable = Nothing

    '    Try

    '        Dim pParam(3) As Object
    '        pParam(0) = UserWS
    '        pParam(1) = PassWS
    '        pParam(2) = Address
    '        pParam(3) = PostalCode

    '        Dim ws As New ServiceDeliveryMan.ServiceDeliveryMan
    '        Dim rResult As Object() = ws.GetLatLonByAddressPostalCode(pParam)

    '        If rResult(0) = "0" Then

    '            Dim result2 As String = "" & rResult(2)

    '            Dim LatLong As String = "Lat,Long"
    '            If result2 <> "" Then
    '                LatLong &= "|" & result2.Replace("|", ",").Replace(" ", "")
    '            End If

    '            dt = ConvertStringToDatatable(LatLong)

    '        Else
    '            ErrorMessage = rResult(1).ToString
    '        End If

    '        Return dt

    '    Catch ex As Exception

    '        ErrorMessage = ex.Message

    '        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
    '        Dim Pesan As String = ""
    '        Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
    '        WriteTracelogTxt(Pesan)

    '        Return dt

    '    Finally

    '        dt = Nothing
    '        GC.Collect()

    '    End Try

    'End Function

    Public Function DashboardPartnerPickupInquiry(ByVal TglAwal As String, ByVal TglAkhir As String, ByVal AccPartner As String, ByVal Status As String, ByVal HubCode As String, ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(6)
            param(0) = ClsWebVer.UserWS 'User
            param(1) = ClsWebVer.PassWS 'Password
            param(2) = TglAwal 'TglAwal 2021-11-25
            param(3) = TglAkhir 'TglAkhir 2021-12-03
            param(4) = AccPartner 'AccPartner
            param(5) = Status 'Status 0: penugasan jemput, 1: konfirmasi jemput, 2: selesai jemput
            param(6) = HubCode 'Untuk Status 1 dan 2, perlu Hub Code, Status 0 tidak perlu

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.DashboardPartnerPickupInquiry(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                'dt = dsData.Tables(i).Copy
                                dt = dsData.Tables(i)
                                dsData.Tables.Remove(dt)
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

            Return dt

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function DashboardPartnerPickupUpdateStatus(ByVal TempUser As String, ByVal Status As String, ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon(3) As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS 'User
            param(1) = ClsWebVer.PassWS 'Password
            param(2) = TempUser 'TempUser
            param(3) = Status 'TempStatus 0: penugasan jemput, 1: konfirmasi jemput, 2: selesai jemput

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.DashboardPartnerPickupUpdateStatus(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                'dt = dsData.Tables(i).Copy
                                dt = dsData.Tables(i)
                                dsData.Tables.Remove(dt)
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

            Return dt

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    'Public Function DashboardPartnerPickupGetSlot() As DataTable

    '    Dim dt As New DataTable

    '    Dim serv As New CoreService.Service
    '    Dim param() As Object
    '    Dim respon() As Object

    '    ReDim param(1)
    '    param(0) = UserWS
    '    param(1) = PassWS

    '    respon = Nothing

    '    Dim i As Integer = 1
    '    Dim success As Boolean = False
    '    While i <= maxTryWS And success = False

    '        Try
    '            respon = serv.DashboardPartnerPickupGetSlot(AppName, AppVersion, param)
    '            success = True

    '        Catch ex As Exception

    '            If i >= maxTryWS Then
    '                Throw
    '            End If

    '            i = i + 1

    '        End Try

    '    End While

    '    If respon(0) = "0" Then

    '        dt = ConvertStringToDatatableWithBarcode(respon(2))

    '    Else

    '        Return Nothing

    '    End If

    '    Return dt

    'End Function

    Public Function DashboardPartnerUpdatePickupRequestAddress(
      ByVal RequestId As String, ByVal UserId As String, ByVal AddressNew As String, ByVal KodePosNew As String _
    , ByVal NamaPicNew As String, ByVal TelpPicNew As String, ByVal ArmadaNew As String _
    , ByVal TanggalPickupNew As String, ByVal JamPickupNew As String, ByVal JamTutupNew As String _
    , ByVal KeteranganNew As String, ByVal AlasanPerubahan As String, ByVal KeteranganPerubahan As String, ByVal Status As String _
    , ByRef ErrorMessage As String) As Boolean

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(15)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = RequestId
            param(3) = UserId
            param(4) = AddressNew
            param(5) = KodePosNew
            param(6) = NamaPicNew
            param(7) = TelpPicNew
            param(8) = ArmadaNew
            param(9) = TanggalPickupNew
            param(10) = JamPickupNew
            param(11) = JamTutupNew
            param(12) = KeteranganNew
            param(13) = AlasanPerubahan
            param(14) = KeteranganPerubahan
            param(15) = Status

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.DashboardPartnerUpdatePickupRequestAddress(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then
                Return True
            Else
                ErrorMessage = respon(1)
            End If

        Catch ex As Exception

            ErrorMessage = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return False

    End Function

    Public Function GetPJBDetailItemByRequestId(ByVal RequestId As String) As DataTable

        Dim SqlQuery As String = ""
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing
        Try
            SqlQuery = "call `sp_dashboardpartnerpickuprequest_detailitem`(@RequestId);"
            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If
            Dim SqlParam As New Dictionary(Of String, String)
            SqlParam.Add("@RequestId", RequestId)
            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If

        End Try

    End Function

    Public Function DashboardPartnerUpdateDetailPickupItems(ByVal RequestId As String, ByVal UserId As String, ByVal DeleteReqItemId As String, ByVal DeleteReqItemName As String, ByRef DsData As DataSet, ByRef PesanError As String) As Boolean

        Try
            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = RequestId
            param(3) = UserId
            param(4) = DeleteReqItemId
            param(5) = DeleteReqItemName

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.DashboardPartnerUpdateDetailPickupItems(ClsWebVer.AppName, ClsWebVer.AppVersion, param, DsData)

            If respon(0) = "0" Then
                Return True
            Else
                PesanError = respon(1)
            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return False

    End Function

    Public Function SetAutoPickupRequestSetting(ByVal HubCode As String, ByVal PickupDateList As String, ByVal UserId As String, ByRef PesanError As String) As Boolean

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(4)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = HubCode
            param(3) = PickupDateList
            param(4) = UserId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.SetAutoPickupRequestSetting(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then
                Return True
            Else
                PesanError = respon(1)
            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return False

    End Function

    Public Function GetHubProcessTypeExpedition(ByVal HubCode As String, ByVal ProcessType As String, ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = HubCode
            param(3) = ProcessType 'SERAH, TERIMA, JEMPUT REKANAN, JEMPUT EKSPEDISI

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetHubProcessTypeExpedition(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                'dt = dsData.Tables(i).Copy
                                dt = dsData.Tables(i)
                                dsData.Tables.Remove(dt)
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

            Return dt

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetHubProcessTypeExpedition(ByVal HubCode As String, ByVal ProcessType As String, ByVal ExpeditionList As String, ByVal UserId As String, ByRef PesanError As String) As Boolean

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = HubCode
            param(3) = ProcessType 'SERAH, TERIMA, JEMPUT REKANAN, JEMPUT EKSPEDISI
            param(4) = ExpeditionList 'Account1|Account2|... (hanya yang aktif)
            param(5) = UserId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.SetHubProcessTypeExpedition(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then
                Return True
            Else
                PesanError = respon(1)
            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return False

    End Function

    Public Function GetMandatoryExpedition(ByVal HubCodeAsal As String, ByVal HubCodeTujuan As String, ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = HubCodeAsal
            param(3) = HubCodeTujuan

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetMandatoryExpeditionCons(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                'dt = dsData.Tables(i).Copy
                                dt = dsData.Tables(i)
                                dsData.Tables.Remove(dt)
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

            Return dt

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SetMandatoryExpedition(ByVal HubCodeAsal As String, ByVal HubCodeTujuan As String, ByVal ExpeditionList As String, ByVal UserId As String, ByRef PesanError As String) As Boolean

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = HubCodeAsal
            param(3) = HubCodeTujuan
            param(4) = ExpeditionList 'Account1|Account2|... (hanya yang aktif)
            param(5) = UserId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.SetMandatoryExpeditionCons(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then
                Return True
            Else
                PesanError = respon(1)
            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return False

    End Function

    Public Function GetRegionHubList(ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetRegionHubList(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                dt = dsData.Tables(i).Copy
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return dt

    End Function

    Public Function SetRegionHubList(ByVal UserId As String, ByRef dsData As DataSet, ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = UserId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.SetRegionHubList(ClsWebVer.AppName, ClsWebVer.AppVersion, param, dsData)

            If respon(0) = "0" Then

                If Not IsNothing(dsData) Then
                    If dsData.Tables.Count > 0 Then

                        For i As Integer = 0 To dsData.Tables.Count - 1

                            If dsData.Tables(i).TableName.ToUpper = "DATA" Then

                                dt = dsData.Tables(i).Copy
                                Exit For

                            End If

                        Next

                    End If
                End If

            Else

                PesanError = respon(1)

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return dt

    End Function

    Public Function PrintSuratTugas(ByVal NoSuratTugas As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByVal NIK As String, ByVal NameOfUser As String, ByRef dtCreatedFilePath As DataTable, ByRef PesanError As String) As Boolean

        Dim NamaRpt As String = "SuratTugas.rpt"
        Dim dt As New DataTable("SuratTugas")

        If CurrentWebCode <> "" Then

            Try

                Dim serv As New CoreService.Service
                Dim param() As Object
                Dim respon(3) As Object

                ReDim param(4)
                param(0) = UserWS
                param(1) = PassWS
                param(2) = "HUB"
                param(3) = CurrentWebCode
                param(4) = NoSuratTugas

                serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

                Dim i As Integer = 1
                Dim success As Boolean = False
                While i <= maxTryWS And success = False

                    Try

                        respon = serv.PrintSuratTugas(AppName, AppVersion, param)
                        success = True

                    Catch ex As Exception

                        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                        Dim Pesan As String = ""
                        Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                        WriteTracelogTxt(Pesan)

                        i = i + 1

                    End Try

                End While

                If success Then

                    If respon(0) = "0" Then

                        If PrintSuratTugas_Rangkuman(respon, NoSuratTugas, CurrentWebCode, CurrentWebName, NIK, NameOfUser, dtCreatedFilePath, PesanError) Then
                        Else
                            Return False
                        End If

                        Dim dtRetur As New DataTable
                        Try
                            dtRetur = ConvertStringToDatatable(respon(4))
                            Dim dtReturTemp As New DataTable
                            dtRetur.DefaultView.RowFilter = ""

                            Dim columnNames(2) As String
                            columnNames(0) = "Account"
                            columnNames(1) = "Name"
                            columnNames(2) = "Alias"

                            dtReturTemp = dtRetur.DefaultView.ToTable(True, columnNames)

                            Dim Index As Integer = 0
                            Dim TotalDst As Integer = dtReturTemp.Rows.Count - 1

                            While Index <= TotalDst

                                dtRetur.DefaultView.RowFilter = "Account = '" & dtReturTemp.Rows(Index).Item("Account") & "' and Alias = '" & dtReturTemp.Rows(Index).Item("Alias") & "'"

                                If PrintSuratTugas_Retur(respon, dtRetur.DefaultView.ToTable _
                                   , dtReturTemp.Rows(Index).Item("Account") _
                                   , dtReturTemp.Rows(Index).Item("Alias") _
                                   , NoSuratTugas, CurrentWebCode, CurrentWebName, NIK, NameOfUser _
                                   , dtCreatedFilePath, PesanError) Then
                                Else
                                    Return False
                                End If

                                Index = Index + 1

                            End While

                            Return True

                        Catch ex As Exception

                            PesanError = ex.Message

                            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                            Dim Pesan As String = ""
                            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                            WriteTracelogTxt(Pesan)

                            Return False

                        Finally

                            dtRetur = Nothing
                            GC.Collect()

                        End Try

                    Else

                        PesanError = respon(1)

                        Dim drCreatedFilePath As DataRow = dtCreatedFilePath.NewRow
                        drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                        drCreatedFilePath.Item("Result") = respon(0)
                        drCreatedFilePath.Item("Message") = respon(1)
                        drCreatedFilePath.Item("Path") = ""
                        dtCreatedFilePath.Rows.Add(drCreatedFilePath)

                        Return False

                    End If

                Else

                    PesanError = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                    WriteTracelogTxt(Pesan)

                    Return False

                End If

                Return False

            Catch ex As Exception

                PesanError = ex.Message
                WriteTracelogTxt(ex.Message)
                Return False

            Finally

                dt = Nothing
                GC.Collect()

            End Try

        End If

    End Function

    Private Function PrintSuratTugas_Rangkuman(ByVal respon() As Object, ByVal NoSuratTugas As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByVal NIK As String, ByVal NameOfUser As String, ByRef dtCreatedFilePath As DataTable, ByRef PesanError As String) As Boolean

        Dim NamaRpt As String = "SuratTugas.rpt"
        Dim dt As New DataTable("SuratTugas")

        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
        Dim FullPathReportFile As String = PathReportFile & "\SuratTugas_" & CurrentWebCode & "_NoST_" & NoSuratTugas & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ".pdf"
        FullPathReportFile = FullPathReportFile.Replace("\\", "\")

        Dim drCreatedFilePath As DataRow

        Try

            Dim header() As String = Nothing
            header = respon(1).ToString.Split("|")

            Dim respCompanyName As String = header(0)
            Dim respCompanyAddress As String = header(1)
            Dim respCompanyPhone As String = header(2)
            Dim respCompanyFax As String = header(3)
            Dim StNum As String = header(4)
            Dim DriverID As String = header(5)
            Dim DriverName As String = header(6)
            Dim Vehicle As String = header(7)
            Dim Expedition As String = header(8)
            Dim Type As String = header(9)
            Dim Code As String = header(10)
            Dim CreatedDate As String = header(11)

            Dim SenderID As String = ""
            Dim SenderName As String = ""
            Try
                SenderID = header(12)
                SenderName = header(13)
            Catch ex As Exception
                SenderID = NIK
                SenderName = NameOfUser
            End Try

            Dim bStNum As String = respon(3).ToString()

            Dim CompanyName As String = respCompanyName
            Dim CompanyAddress As String = respCompanyAddress

            Dim CompanyPhone As String = ""
            If respCompanyPhone.Trim <> "" Then
                CompanyPhone = "Phone : " & respCompanyPhone
            End If

            Dim CompanyFax As String = ""
            If respCompanyFax.Trim <> "" Then
                CompanyFax = "Fax : " & respCompanyFax
            End If

            Dim DocumentName As String = "SURAT TUGAS"

            Dim DocumentNo As String = "NO. " & StNum & " / " & Code
            If CurrentWebName <> "" Then
                DocumentNo &= "-" & CurrentWebName
            End If

            Dim PetugasStationName As String = SenderName
            Dim PetugasStationID As String = SenderID
            Dim BarcodeDocumentNo As String = bStNum
            Dim TanggalDocument As String = CreatedDate
            Dim ExpeditionName As String = Expedition
            Dim VehicleNo As String = Vehicle
            Dim Origin As String = Code

            If CurrentWebName <> "" Then
                Origin &= "-" & CurrentWebName
            End If

            dt = ConvertStringToDatatable(respon(2))

            Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

            rpt.Load(FileRptPath)
            rpt.SummaryInfo.ReportTitle = "Surat Tugas"
            rpt.PrintOptions.PaperSize = PaperSize.PaperA4
            rpt.SetDataSource(dt)

            'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
            'Dim crParameterFieldLocation As ParameterFieldDefinition
            Dim crParameterValues As ParameterValues
            Dim crParameterDiscreteValue As ParameterDiscreteValue

            Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyAddress")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyAddress")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyAddress
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyPhone")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyPhone")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyPhone
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyFax")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyFax")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyFax
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = BarcodeDocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TanggalDocument")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalDocument")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = TanggalDocument
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ExpeditionName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ExpeditionName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ExpeditionName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("VehicleNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("VehicleNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = VehicleNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Origin")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("Origin")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = Origin
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PetugasStationName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasStationName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = PetugasStationName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PetugasStationID")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasStationID")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = PetugasStationID
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DriverName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverID")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverID")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DriverID
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                Dim ShowLogo As String = ReadWebConfig("RPTShowLogo", False)
                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ShowLogo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ShowLogo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ShowLogo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

            End Using

            'Pindah ke paling atas
            'Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
            'Dim FullPathReportFile As String = PathReportFile & "\SuratTugas_" & CurrentWebCode & "_NoST_" & NoSuratTugas & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".pdf"
            'FullPathReportFile = FullPathReportFile.Replace("\\", "\")

            'Dim drCreatedFilePath As DataRow

            Try
                Dim PrintToPDF As String = "1" 'ReadWebConfig("PrintToPDF", False)
                If PrintToPDF = "1" Then

                    If Not Directory.Exists(PathReportFile) Then
                        Directory.CreateDirectory(PathReportFile)
                    End If

                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)

                    drCreatedFilePath = dtCreatedFilePath.NewRow
                    drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                    drCreatedFilePath.Item("Result") = "0"
                    drCreatedFilePath.Item("Message") = "Berhasil"
                    drCreatedFilePath.Item("Path") = FullPathReportFile
                    dtCreatedFilePath.Rows.Add(drCreatedFilePath)

                End If
            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                drCreatedFilePath = dtCreatedFilePath.NewRow
                drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                drCreatedFilePath.Item("Result") = "9"
                drCreatedFilePath.Item("Message") = ex.Message
                drCreatedFilePath.Item("Path") = FullPathReportFile
                dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            End Try

            'Try
            '    crParameterFieldDefinitions.Dispose()
            'Catch ex As Exception
            'End Try

            'Try
            '    crParameterFieldLocation.Dispose()
            'Catch ex As Exception
            'End Try

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            drCreatedFilePath = dtCreatedFilePath.NewRow
            drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
            drCreatedFilePath.Item("Result") = "9"
            drCreatedFilePath.Item("Message") = ex.Message
            drCreatedFilePath.Item("Path") = FullPathReportFile
            dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            Return False

        Finally

            Try
                rpt.Close()
            Catch ex As Exception
            End Try
            Try
                rpt.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            drCreatedFilePath = Nothing

            GC.Collect()

        End Try

    End Function

    Private Function PrintSuratTugas_Retur(ByVal respon() As Object, ByVal dt_respon2 As DataTable, ByVal Account As String, ByVal NamaEComm As String, ByVal NoSuratTugas As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByVal NIK As String, ByVal NameOfUser As String, ByRef dtCreatedFilePath As DataTable, ByRef PesanError As String) As Boolean

        Dim NamaRpt As String = "SuratTugasRetur.rpt"
        Dim dt As New DataTable("SuratTugasRetur")

        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
        Dim FullPathReportFile As String = PathReportFile & "\SuratTugas_" & CurrentWebCode & "_NoST_" & NoSuratTugas & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & "_RTR_" & Account & ".pdf"
        FullPathReportFile = FullPathReportFile.Replace("\\", "\")

        Dim drCreatedFilePath As DataRow

        Try

            Dim header() As String = Nothing
            header = respon(1).ToString.Split("|")

            Dim respCompanyName As String = header(0)
            Dim respCompanyAddress As String = header(1)
            Dim respCompanyPhone As String = header(2)
            Dim respCompanyFax As String = header(3)
            Dim StNum As String = header(4)
            Dim DriverID As String = header(5)
            Dim DriverName As String = header(6)
            Dim Vehicle As String = header(7)
            Dim Expedition As String = header(8)
            Dim Type As String = header(9)
            Dim Code As String = header(10)
            Dim CreatedDate As String = header(11)

            Dim SenderID As String = ""
            Dim SenderName As String = ""
            Try
                SenderID = header(12)
                SenderName = header(13)
            Catch ex As Exception
                SenderID = NIK
                SenderName = NameOfUser
            End Try

            Dim bStNum As String = respon(3).ToString()

            Dim CompanyName As String = respCompanyName
            Dim CompanyAddress As String = respCompanyAddress

            Dim CompanyPhone As String = ""
            If respCompanyPhone.Trim <> "" Then
                CompanyPhone = "Phone : " & respCompanyPhone
            End If

            Dim CompanyFax As String = ""
            If respCompanyFax.Trim <> "" Then
                CompanyFax = "Fax : " & respCompanyFax
            End If

            Dim DocumentName As String = "SURAT TUGAS (AWB RETUR)"

            Dim DocumentNo As String = "NO. " & StNum & " / " & Code
            If CurrentWebName <> "" Then
                DocumentNo &= "-" & CurrentWebName
            End If

            Dim PetugasStationName As String = SenderName
            Dim PetugasStationID As String = SenderID
            Dim BarcodeDocumentNo As String = bStNum
            Dim TanggalDocument As String = CreatedDate
            Dim ExpeditionName As String = Expedition
            Dim VehicleNo As String = Vehicle

            Dim Origin As String = Code
            If CurrentWebName <> "" Then
                Origin &= "-" & CurrentWebName
            End If

            dt = dt_respon2

            Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

            rpt.Load(FileRptPath)
            rpt.SummaryInfo.ReportTitle = "Surat Tugas Retur"
            rpt.PrintOptions.PaperSize = PaperSize.PaperA4
            rpt.SetDataSource(dt)

            'Dim crParameterFieldDefinitions As ParameterFieldDefinitions
            'Dim crParameterFieldLocation As ParameterFieldDefinition
            Dim crParameterValues As ParameterValues
            Dim crParameterDiscreteValue As ParameterDiscreteValue

            Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyAddress")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyAddress")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyAddress
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyPhone")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyPhone")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyPhone
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyFax")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyFax")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyFax
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = BarcodeDocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TanggalDocument")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalDocument")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = TanggalDocument
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ExpeditionName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ExpeditionName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ExpeditionName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("VehicleNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("VehicleNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = VehicleNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("Origin")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("Origin")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = Origin
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PetugasStationName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasStationName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = PetugasStationName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("PetugasStationID")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasStationID")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = PetugasStationID
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DriverName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DriverID")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DriverID")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DriverID
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("NamaEComm")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("NamaEComm")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = NamaEComm
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                Dim ShowLogo As String = ReadWebConfig("RPTShowLogo", False)
                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ShowLogo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ShowLogo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ShowLogo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

            End Using

            'Pindah ke paling atas
            'Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
            'Dim FullPathReportFile As String = PathReportFile & "\SuratTugas_" & CurrentWebCode & "_NoST_" & NoSuratTugas & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmss") & "_RTR_" & Account & ".pdf"
            'FullPathReportFile = FullPathReportFile.Replace("\\", "\")

            'Dim drCreatedFilePath As DataRow

            Try
                Dim PrintToPDF As String = "1" 'ReadWebConfig("PrintToPDF", False)
                If PrintToPDF = "1" Then

                    If Not Directory.Exists(PathReportFile) Then
                        Directory.CreateDirectory(PathReportFile)
                    End If

                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)

                    drCreatedFilePath = dtCreatedFilePath.NewRow
                    drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                    drCreatedFilePath.Item("Result") = "0"
                    drCreatedFilePath.Item("Message") = "Berhasil"
                    drCreatedFilePath.Item("Path") = FullPathReportFile
                    dtCreatedFilePath.Rows.Add(drCreatedFilePath)

                End If
            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                drCreatedFilePath = dtCreatedFilePath.NewRow
                drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                drCreatedFilePath.Item("Result") = "9"
                drCreatedFilePath.Item("Message") = ex.Message
                drCreatedFilePath.Item("Path") = FullPathReportFile
                dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            End Try

            'Try
            '    crParameterFieldDefinitions.Dispose()
            'Catch ex As Exception
            'End Try

            'Try
            '    crParameterFieldLocation.Dispose()
            'Catch ex As Exception
            'End Try

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            drCreatedFilePath = dtCreatedFilePath.NewRow
            drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
            drCreatedFilePath.Item("Result") = "9"
            drCreatedFilePath.Item("Message") = ex.Message
            drCreatedFilePath.Item("Path") = FullPathReportFile
            dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            Return False

        Finally

            Try
                rpt.Close()
            Catch ex As Exception
            End Try
            Try
                rpt.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            drCreatedFilePath = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function PrintSuratJemputBarang(ByVal RequestId As String, ByVal CurrentWebCode As String, ByVal NIK As String, ByVal DocumentName As String, ByRef dtCreatedFilePath As DataTable, ByRef PesanError As String) As Boolean

        Dim NamaRpt As String = "SuratTugas.rpt"
        Dim dt As New DataTable("SuratTugas")

        If CurrentWebCode <> "" Then

            Try

                Dim serv As New CoreService.Service
                Dim param() As Object
                Dim respon(3) As Object

                ReDim param(5)
                param(0) = UserWS
                param(1) = PassWS
                param(2) = DocumentName
                param(3) = "PJB"
                param(4) = CurrentWebCode
                param(5) = RequestId

                serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

                Dim i As Integer = 1
                Dim success As Boolean = False
                While i <= maxTryWS And success = False
                    Try

                        respon = serv.PrintSuratJemputBarang(AppName, AppVersion, param)
                        success = True

                    Catch ex As Exception

                        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                        Dim Pesan As String = ""
                        Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                        WriteTracelogTxt(Pesan)

                        i = i + 1

                    End Try

                End While

                If success Then

                    If respon(0) = "0" Then

                        If Not PrintSuratTugas_SuratJemputBarang(respon, RequestId, CurrentWebCode, NIK, dtCreatedFilePath, PesanError) Then
                            Return False
                        End If

                    Else

                        PesanError = respon(1)

                        Dim drCreatedFilePath As DataRow = dtCreatedFilePath.NewRow
                        drCreatedFilePath.Item("NoSuratTugas") = RequestId
                        drCreatedFilePath.Item("Result") = respon(0)
                        drCreatedFilePath.Item("Message") = respon(1)
                        drCreatedFilePath.Item("Path") = ""
                        dtCreatedFilePath.Rows.Add(drCreatedFilePath)

                        Return False

                    End If

                Else

                    PesanError = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                    WriteTracelogTxt(Pesan)

                    Return False

                End If

                Return False

            Catch ex As Exception

                PesanError = ex.Message
                WriteTracelogTxt(ex.Message)
                Return False

            Finally

                dt = Nothing
                GC.Collect()

            End Try

        End If

    End Function

    Private Function PrintSuratTugas_SuratJemputBarang(ByVal respon() As Object, ByVal NoSuratTugas As String, ByVal CurrentWebCode As String, ByVal NIK As String, ByRef dtCreatedFilePath As DataTable, ByRef PesanError As String) As Boolean

        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim NamaRpt As String = "SuratTugasPJB.rpt"
        Dim dt As New DataTable("SuratTugas")

        Dim PathReportFile As String = ReadWebConfig("PathReportFilePJB", False)
        Dim FullPathReportFile As String = ""

        Dim drCreatedFilePath As DataRow
        Try

            Dim header() As String = Nothing
            header = respon(1).ToString.Split("|")

            Dim DocumentName As String = header(0)
            Dim respCompanyName As String = header(1)
            Dim respCompanyAddress As String = header(2)
            Dim respCompanyPhone As String = header(3)
            Dim respCompanyFax As String = header(4)
            Dim StNum As String = header(5)
            Dim Expedition As String = header(6)
            Dim Code As String = header(7)
            Dim CreatedDate As String = header(8)

            Dim bStNum As String = respon(3).ToString()

            Dim CompanyName As String = respCompanyName
            Dim CompanyAddress As String = respCompanyAddress

            Dim CompanyPhone As String = ""
            If respCompanyPhone.Trim <> "" Then
                CompanyPhone = "Phone : " & respCompanyPhone
            End If

            Dim CompanyFax As String = ""
            If respCompanyFax.Trim <> "" Then
                CompanyFax = "Fax : " & respCompanyFax
            End If

            Dim DocumentNo As String = "NO. " & StNum & " / " & Code

            Dim BarcodeDocumentNo As String = bStNum
            Dim TanggalDocument As String = CreatedDate
            Dim ExpeditionName As String = Expedition

            dt = ConvertStringToDatatable(respon(2))

            Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

            rpt.Load(FileRptPath)
            rpt.SummaryInfo.ReportTitle = "SuratTugasPJB"
            rpt.PrintOptions.PaperSize = PaperSize.PaperA4
            rpt.SetDataSource(dt)


            Dim crParameterValues As ParameterValues
            Dim crParameterDiscreteValue As ParameterDiscreteValue

            Using crParameterFieldDefinitions As ParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyAddress")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyAddress")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyAddress
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyPhone")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyPhone")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyPhone
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("CompanyFax")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("CompanyFax")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = CompanyFax
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("DocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("BarcodeDocumentNo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = BarcodeDocumentNo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("TanggalDocument")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalDocument")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = TanggalDocument
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ExpeditionName")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ExpeditionName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ExpeditionName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using


                Dim ShowLogo As String = ReadWebConfig("RPTShowLogo", False)
                'crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
                Using crParameterFieldLocation As ParameterFieldDefinition = crParameterFieldDefinitions.Item("ShowLogo")
                    'crParameterFieldLocation = crParameterFieldDefinitions.Item("ShowLogo")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = ShowLogo
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)
                End Using

            End Using

            FullPathReportFile &= PathReportFile & "\"
            If DocumentName.Contains("REPRINT") Then
                FullPathReportFile &= "Reprint"
            End If
            FullPathReportFile &= "PJB_" & CurrentWebCode & "_NoST_" & NoSuratTugas & "_NIK_" & NIK & ".pdf"
            FullPathReportFile = FullPathReportFile.Replace("\\", "\")

            Try
                Dim PrintToPDF As String = "1" 'ReadWebConfig("PrintToPDF", False)
                If PrintToPDF = "1" Then

                    If Not Directory.Exists(PathReportFile) Then
                        Directory.CreateDirectory(PathReportFile)
                    End If

                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)

                    drCreatedFilePath = dtCreatedFilePath.NewRow
                    drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                    drCreatedFilePath.Item("Result") = "0"
                    drCreatedFilePath.Item("Message") = "Berhasil"
                    drCreatedFilePath.Item("Path") = FullPathReportFile
                    dtCreatedFilePath.Rows.Add(drCreatedFilePath)

                End If
            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                drCreatedFilePath = dtCreatedFilePath.NewRow
                drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
                drCreatedFilePath.Item("Result") = "9"
                drCreatedFilePath.Item("Message") = ex.Message
                drCreatedFilePath.Item("Path") = FullPathReportFile
                dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            End Try
            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            drCreatedFilePath = dtCreatedFilePath.NewRow
            drCreatedFilePath.Item("NoSuratTugas") = NoSuratTugas
            drCreatedFilePath.Item("Result") = "9"
            drCreatedFilePath.Item("Message") = ex.Message
            drCreatedFilePath.Item("Path") = FullPathReportFile
            dtCreatedFilePath.Rows.Add(drCreatedFilePath)

            Return False

        Finally

            Try
                rpt.Close()
            Catch ex As Exception
            End Try
            Try
                rpt.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            drCreatedFilePath = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function GetToolsAWBGroupList() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetToolsAwbGroupList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetToolsAwbGroup() As DataTable

        Dim dt As DataTable = Nothing

        Dim errmsg As String = ""

        Try
            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("coreserviceurl")
            respon = serv.GetToolsAwbGroupAll(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "data"

            Else

                errmsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            errmsg = ex.Message

            Dim methodname As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim pesan As String = ""
            pesan = "dari clsfungsi, proses " & methodname & ", error : " & ex.Message
            WriteTracelogTxt(pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If errmsg <> "" Then
                dt = New DataTable
                dt.TableName = "error"
                dt.Columns.Add("respon")
                dt.Rows.Add(errmsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function


#Region "Partner Item"

    Public Function PartnerItemListReportCategoryNoFilter() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.PartnerItemListReportCategory(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatableWithBarcode(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            'WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function PartnerItemListItemNoFilter() As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.PartnerItemListAllItem(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            'WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

#End Region


    Public Function GetFinFinalProcess(ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing
        Dim MCon As MySqlConnection = Nothing

        Try

            Dim SqlQuery As String = ""
            SqlQuery = "select `Alias` as `Display`, `Type` as `Value`"
            SqlQuery &= " from `accountfintype`"
            SqlQuery &= " where curdate() between activedate and ifnull(inactivedate, curdate())"
            SqlQuery &= ";"

            Dim SqlParam As New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetHubPettyCashCutOffList(ByRef ErrorMessage As String) As DataTable

        Dim SqlQuery As String = ""
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = ""
            SqlQuery &= " Select Hub"
            SqlQuery &= "  , HubName"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashDate', `Value`, '2000-12-31')) as char) as TglCutOff"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashDate', `Value`, '2000-12-31')) as char) as CurrTglCutOff"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashSaldo', `Value`, '0')) as char) as SaldoAwal"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashSaldo', `Value`, '0')) as char) as CurrSaldoAwal"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashLstPst', `Value`, '2000-12-31')) as char) as LastPosting"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashLstPst', `Value`, '2000-12-31')) as char) as CurrLastPosting"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashLiFiTu', `Value`, '0')) as char) as LiFiTu"
            SqlQuery &= "  , cast(MAX(IF(Key1 = 'PettyCashLiFiTu', `Value`, '0')) as char) as CurrLiFiTu"
            SqlQuery &= " from ("
            SqlQuery &= "   select c.Key2 as Hub"
            SqlQuery &= "    , cast(if(h.Hub is not null, concat(h.Alias,' - ',h.Hub), h.Hub) as char) as HubName"
            SqlQuery &= "    , Key1"
            SqlQuery &= "    , `Value` as `Value`"
            SqlQuery &= "   from const c"
            SqlQuery &= "   left join msthub h on h.Hub = c.Key2"
            SqlQuery &= "   where key1 in ('PettyCashDate','PettyCashSaldo','PettyCashLstPst','PettyCashLiFiTu')"
            SqlQuery &= " ) tbl1 group by Hub"
            SqlQuery &= " order by HubName"
            SqlQuery &= " ;"

            Dim SqlParam As New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function GetAccountCategory(ByRef ErrorMessage As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "select cast('' as char) as `Display`, cast('' as char) as `Value` union "
            SqlQuery &= "select `Description` as `Display`, `Category` as `Value`"
            SqlQuery &= " from accountcategory"
            SqlQuery &= " where curdate() between activedate and ifnull(inactivedate, curdate())"
            SqlQuery &= " ;"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function GetAccountSubCategory(ByRef ErrorMessage As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "select cast('' as char) as `Display`, cast('' as char) as `Value`, cast('' as char) as `MappingCategory` union "
            SqlQuery &= "select `Description` as `Display`, `SubCategory` as `Value`, `Category` as `MappingCategory`"
            SqlQuery &= " from accountsubcategory"
            SqlQuery &= " where curdate() between activedate and ifnull(inactivedate, curdate())"
            SqlQuery &= " ;"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function Get_GroupAccess(ByVal Code As String, ByVal UserName As String) As String

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing
        Dim result As String = ""

        Try

            SqlQuery = "call `sp_get_webreport_groupaccess`(@Code,@UserName);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Code", Code)
            SqlParam.Add("@UserName", UserName)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    result = "" & dt.Rows(0).Item(0)
                End If
            End If

            Return result

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return result

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function getKeluhanStatusOPR(ByRef PesanError As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim dt As DataTable = Nothing
        Dim MCon As MySqlConnection = Nothing

        Try

            'complainstatusopr
            SqlQuery = "select * from (Select '' as `Display`, '' as `Value`, '' as `MustFillSolution`) t1 union"
            SqlQuery &= " select * from (Select `Name` as `Display`, `ID` as `Value`, `MustFillSolution` from complainstatusopr order by seqno) t2"
            SqlQuery &= ";"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            PesanError = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function


    Public Function getDataCustomRateRequest(ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `sp_customrate_request_list`();"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function UpdateDataCustomRateRequest(ByVal NomorRequest As String, ByVal Tipe As String, ByVal ActiveDate As String, ByVal FileName As String, ByVal NIK As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = " update webreport_customrate_request "
            SqlQuery &= " set "
            SqlQuery &= " Pricing_ProcessedDate = NOW() "
            SqlQuery &= " , Pricing_CustomType = @Tipe "
            If ("" & ActiveDate).ToString <> "" Then
                SqlQuery &= " , Pricing_ActiveDate = @ActiveDate "
                SqlParam.Add("@ActiveDate", ActiveDate)
            Else
                SqlQuery &= " , Pricing_ActiveDate = curdate() "
            End If
            SqlQuery &= " , Pricing_Filename = @FileName "
            SqlQuery &= " , Pricing_UserID = @NIK "
            SqlQuery &= " where true "
            SqlQuery &= " and NomorREQ = @NomorRequest "
            SqlQuery &= ";"

            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@FileName", FileName)
            SqlParam.Add("@NIK", NIK)
            SqlParam.Add("@NomorRequest", NomorRequest)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Return ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            SqlParam = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function SaveCustomRateRequest(ByVal Account As String, ByVal Tipe As String, ByVal TglAktif As String, ByVal NamaFile As String, ByVal UserID As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `insert_customrate`('DIRECT',@Account,NULL,NULL,'','','',NOW(),@Tipe,@TglAktif,@NamaFile,@UserID);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Account", Account)
            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@TglAktif", TglAktif)
            SqlParam.Add("@NamaFile", NamaFile)
            SqlParam.Add("@UserID", UserID)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Dim Result As String = ObjSQL.SQLExecuteScalar(MCon, SqlQuery, SqlParam)

            If Result.ToUpper = "SUCCESS" Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            SqlParam = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function getDataCustomRateSellerCategoryRequest(ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `sp_customrate_sellercategory_request_list`();"

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            ErrorMessage = ex.Message

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If
        End Try

        Return dt

    End Function

    Public Function UpdateDataCustomRateSellerCategoryRequest(ByVal NomorRequest As String, ByVal Tipe As String, ByVal ActiveDate As String, ByVal FileName As String, ByVal NIK As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = " update webreport_customrate_sellercategory_request "
            SqlQuery &= " set "
            SqlQuery &= " Pricing_ProcessedDate = NOW() "
            SqlQuery &= " , Pricing_CustomType = @Tipe "
            If ("" & ActiveDate).ToString <> "" Then
                SqlQuery &= " , Pricing_ActiveDate = @ActiveDate "
                SqlParam.Add("@ActiveDate", ActiveDate)
            Else
                SqlQuery &= " , Pricing_ActiveDate = curdate() "
            End If
            SqlQuery &= " , Pricing_Filename = @FileName "
            SqlQuery &= " , Pricing_UserID = @NIK "
            SqlQuery &= " where true "
            SqlQuery &= " and NomorREQ = @NomorRequest "
            SqlQuery &= ";"

            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@FileName", FileName.Replace("'", "''"))
            SqlParam.Add("@NIK", NIK)
            SqlParam.Add("@NomorRequest", NomorRequest)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Return ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If
        End Try

        Return False

    End Function

    Public Function SaveCustomRateSellerCategoryRequest(ByVal Category As String, ByVal Tipe As String, ByVal TglAktif As String, ByVal NamaFile As String, ByVal UserID As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `insert_customrate_sellercategory`('DIRECT',@Category,NULL,NULL,'','',NOW(),@Tipe,@TglAktif,@NamaFile,@UserID);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Category", Category)
            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@TglAktif", TglAktif)
            SqlParam.Add("@NamaFile", NamaFile)
            SqlParam.Add("@UserID", UserID)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Dim Result As String = ObjSQL.SQLExecuteScalar(MCon, SqlQuery, SqlParam)

            If Result.ToUpper = "SUCCESS" Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            SqlParam = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function getDataCustomRateSellerSubCategoryRequest(ByRef ErrorMessage As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `sp_customrate_sellerSubCategory_request_list`();"

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            ErrorMessage = ex.Message

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If
        End Try

        Return dt

    End Function

    Public Function UpdateDataCustomRateSellerSubCategoryRequest(ByVal NomorRequest As String, ByVal Tipe As String, ByVal ActiveDate As String, ByVal FileName As String, ByVal NIK As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = " update webreport_customrate_sellerSubCategory_request "
            SqlQuery &= " set "
            SqlQuery &= " Pricing_ProcessedDate = NOW() "
            SqlQuery &= " , Pricing_CustomType = @Tipe "
            If ("" & ActiveDate).ToString <> "" Then
                SqlQuery &= " , Pricing_ActiveDate = @ActiveDate "
                SqlParam.Add("@ActiveDate", ActiveDate)
            Else
                SqlQuery &= " , Pricing_ActiveDate = curdate() "
            End If
            SqlQuery &= " , Pricing_Filename = @FileName "
            SqlQuery &= " , Pricing_UserID = @NIK "
            SqlQuery &= " where true "
            SqlQuery &= " and NomorREQ = @NomorRequest "
            SqlQuery &= ";"

            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@FileName", FileName.Replace("'", "''"))
            SqlParam.Add("@NIK", NIK)
            SqlParam.Add("@NomorRequest", NomorRequest)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Return ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If
        End Try

        Return False

    End Function

    Public Function SaveCustomRateSellerSubCategoryRequest(ByVal SubCategory As String, ByVal Tipe As String, ByVal TglAktif As String, ByVal NamaFile As String, ByVal UserID As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            SqlQuery = "call `insert_customrate_sellerSubCategory`('DIRECT',@SubCategory,NULL,NULL,'','',NOW(),@Tipe,@TglAktif,@NamaFile,@UserID);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@SubCategory", SubCategory)
            SqlParam.Add("@Tipe", Tipe)
            SqlParam.Add("@TglAktif", TglAktif)
            SqlParam.Add("@NamaFile", NamaFile)
            SqlParam.Add("@UserID", UserID)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Dim Result As String = ObjSQL.SQLExecuteScalar(MCon, SqlQuery, SqlParam)

            If Result.ToUpper = "SUCCESS" Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            SqlParam = Nothing
            GC.Collect()

        End Try

    End Function


    Public Function InsertMulaiNomorTugas(ByVal Deskripsi As String, ByVal IdTugas As String, ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByVal Creator As String, ByRef NomorTugas As String, ByRef ErrorMessage As String) As Boolean

        If Not IsNothing(Creator) Then

            Dim dt As DataTable = Nothing
            Dim SqlQuery As String = ""
            Dim SqlParam As New Dictionary(Of String, String)
            Dim MCon As MySqlConnection = Nothing

            Try

                SqlQuery = "call insert_pricing_tugas(@Deskripsi,@IdTugas,@TanggalAwal,@TanggalAkhir,@Creator)"

                SqlParam = New Dictionary(Of String, String)
                SqlParam.Add("@Deskripsi", Deskripsi)
                SqlParam.Add("@IdTugas", IdTugas)
                SqlParam.Add("@TanggalAwal", TanggalAwal)
                SqlParam.Add("@TanggalAkhir", TanggalAkhir)
                SqlParam.Add("@Creator", Creator)

                MCon = ObjCon.SetConn_Master
                If MCon.State <> ConnectionState.Open Then
                    MCon.Open()
                End If

                dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

                If Not IsNothing(dt) Then
                    If dt.Rows.Count > 0 Then
                        NomorTugas = dt.Rows(0).Item("result_value")
                        Return True
                    Else
                        ErrorMessage = "Failed - Tidak ada baris pada respon!"
                        Return False
                    End If
                Else
                    Return False
                End If

            Catch ex As Exception

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
                WriteTracelogTxt(Pesan)

                ErrorMessage = ex.Message

                Return False

            Finally

                Try
                    MCon.Close()
                Catch ex As Exception
                End Try

                Try
                    MCon.Dispose()
                Catch ex As Exception
                End Try

                dt = Nothing
                SqlParam = Nothing

                GC.Collect()

            End Try

        Else

            Return False

        End If

    End Function

    Public Function UpdateRTSDate_getDaftar(ByVal NoAWB As String, ByVal UseCurDateNewRTSDate As Boolean, ByRef ErrorMessage As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            Dim tmpTblName As String = "tmpNoAWB"
            Dim Separator As String = "|"

            SqlQuery = "CALL SeparateValues(@TableName,@NoAWB,@Separator)"
            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@TableName", tmpTblName)
            SqlParam.Add("@NoAWB", NoAWB)
            SqlParam.Add("@Separator", Separator)

            ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam)


            SqlQuery = "SELECT tr.TrackNum"
            SqlQuery &= "  , IF(ac.Account IS NOT NULL, CONCAT(ac.Alias,' (',RIGHT(ac.account,3),')'), tr.ShAccount) AS AccountAlias "
            SqlQuery &= "  , tr.CoName AS CoName "
            SqlQuery &= "  , IF(istdst.Code IS NOT NULL, CONCAT(istdst.Name,' (',istdst.Code,')'), tr.DstStore) AS StoreName "
            SqlQuery &= "  , CAST(tpin.RTSDate AS CHAR) AS OldRTSDate "
            If UseCurDateNewRTSDate Then
                SqlQuery &= "  , CAST(CURDATE() AS CHAR(10)) AS NewRTSDate "
            Else
                SqlQuery &= "  , CAST('' AS CHAR(10)) AS NewRTSDate "
            End If
            SqlQuery &= " FROM `transaction` tr "
            SqlQuery &= " INNER JOIN `" & tmpTblName & "` tmp ON tmp.col1 = tr.TrackNum "
            SqlQuery &= " LEFT JOIN account ac ON ac.Account = tr.ShAccount "
            SqlQuery &= " LEFT JOIN indomaretstore istdst ON istdst.Code = tr.DstStore "
            SqlQuery &= " LEFT JOIN transactionpin tpin ON tpin.TrackNum = tr.TrackNum "
            SqlQuery &= " WHERE tpin.PIN IS NOT NULL "

            SqlParam = New Dictionary(Of String, String)

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function UpdateRTSDate_SimpanData(ByVal SimpanData As DataTable, ByVal UserID As String, ByRef ErrorMessage As String) As Boolean

        Dim SqlQuery As String = ""
        Dim SqlParam As Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing

        Try

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            'TrackNum,OldShAccount,NewShAccount,OldCoAccount,NewCoAccount

            Dim PesanError As String = ""

            For i As Integer = 0 To SimpanData.Rows.Count - 1

                PesanError = ""

                SqlQuery = "call `UpdateRTSDate` ("
                SqlQuery &= " @AppName" 'Param AppName
                SqlQuery &= ",@AppVersion" 'Param AppVersion
                SqlQuery &= ",@TrackNum"
                SqlQuery &= ",@OldRTSDate"
                SqlQuery &= ",@NewRTSDate"
                SqlQuery &= ",@UserID"
                SqlQuery &= ")"

                SqlParam = New Dictionary(Of String, String)
                SqlParam.Add("@AppName", AppName)
                SqlParam.Add("@AppVersion", AppVersion)
                SqlParam.Add("@TrackNum", SimpanData.Rows(i).Item("TrackNum"))
                SqlParam.Add("@OldRTSDate", SimpanData.Rows(i).Item("OldRTSDate"))
                SqlParam.Add("@NewRTSDate", SimpanData.Rows(i).Item("NewRTSDate"))
                SqlParam.Add("@UserID", UserID)

                If Not ObjSQL.SQLExecuteNonQuery(MCon, SqlQuery, SqlParam, PesanError) Then

                    Dim exMessage As String = "Error Simpan Data, Pesan Error : " & PesanError & ", Query : " & SqlQuery

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & exMessage
                    WriteTracelogTxt(Pesan)

                    Dim SBError As New StringBuilder

                    SBError.Append("Error Simpan Data : ")
                    SBError.Append("Tracknum = " & SimpanData.Rows(i).Item("TrackNum"))
                    SBError.Append(", Tanggal Retur Lama = " & SimpanData.Rows(i).Item("OldRTSDate"))
                    SBError.Append(", Tanggal Retur Baru = " & SimpanData.Rows(i).Item("NewRTSDate"))

                    ErrorMessage = SBError.ToString

                    SBError = Nothing

                    Return False

                End If

            Next

            Return True

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            SqlParam = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function UpdateRTSDate_KirimEmail(ByVal SimpanData As DataTable, ByRef ErrorMessage As String) As Boolean

        Try

            Dim devWord As String = ReadWebConfig("DevWord")

            Dim _MailTo As String = ReadWebConfig("EmailITToko")
            If _MailTo = "" Then
                ErrorMessage = "Gagal mengirim email, EmailITToko belum disetting!"
                Return False
            End If

            Dim AutoNum As String = Date.Now.ToString("yyyyMMddHHmmss")

            Dim _MailCc As String = "support@indopaket.co.id"
            Dim _Subject As String = devWord & "Indopaket - Permintaan perubahan Tanggal Retur di Toko " & AutoNum
            Dim _Body As String = DatatableToHTML(SimpanData)
            Dim _Filename As String = ""
            Dim _eType As String = ""

            If EmailSend(_MailTo, _MailCc, _Subject, _Body, _Filename, _eType) Then
                Return True
            Else

                Dim exMessage As String = "Error Kirim Email,"
                exMessage &= " Mail To : " & _MailTo
                exMessage &= " Mail Cc : " & _MailCc
                exMessage &= " Subject : " & _Subject

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & ", Error : " & exMessage
                WriteTracelogTxt(Pesan)

                ErrorMessage = exMessage

                Return False

            End If

            Return False

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            ErrorMessage = ex.Message

            Return False

        End Try

    End Function

    Public Function GetDbTableInfo(ByVal TableName As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = TableName

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetDbInfo(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function LimitPeriode(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByRef PesanError As String) As Boolean

        Try

            Dim LimitPeriodeTarikReport As String = ReadWebConfig("LimitPeriodeTarikReport")

            Dim MaxLimitPeriode As Integer = 90
            Try
                MaxLimitPeriode = CType(LimitPeriodeTarikReport, Integer)
            Catch ex As Exception
                MaxLimitPeriode = 90
            End Try

            Dim intYYYY As Integer = 0
            Dim intMM As Integer = 0
            Dim intDD As Integer = 0

            'TanggalAwal Mulai
            Dim dateTanggalAwal As Date = Nothing

            TanggalAwal = TanggalAwal.Trim

            Dim SplitTanggalJamAwal As String() = TanggalAwal.Split(" ") 'Jaga2 jika ada Jam nya
            If SplitTanggalJamAwal.Length < 1 Then
                PesanError = "Format Tanggal Awal yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                Return False
            Else

                Dim SplitTanggalAwal As String() = ("" & SplitTanggalJamAwal(0)).ToString.Trim.Split("-")

                If SplitTanggalAwal.Length < 3 Then
                    PesanError = "Format Tanggal Awal yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                    Return False
                Else

                    intYYYY = CType(("" & SplitTanggalAwal(0)).ToString, Integer)
                    intMM = CType(("" & SplitTanggalAwal(1)).ToString, Integer)
                    intDD = CType(("" & SplitTanggalAwal(2)).ToString, Integer)

                    dateTanggalAwal = New Date(intYYYY, intMM, intDD)

                End If

            End If
            'TanggalAwal Selesai

            'TanggalAkhir Mulai
            Dim dateTanggalAkhir As Date = Nothing

            TanggalAkhir = TanggalAkhir.Trim

            Dim SplitTanggalJamAkhir As String() = TanggalAkhir.Split(" ") 'Jaga2 jika ada Jam nya
            If SplitTanggalJamAkhir.Length < 1 Then
                PesanError = "Format Tanggal Akhir yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                Return False
            Else

                Dim SplitTanggalAkhir As String() = ("" & SplitTanggalJamAkhir(0)).ToString.Trim.Split("-")

                If SplitTanggalAkhir.Length < 3 Then
                    PesanError = "Format Tanggal Akhir yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                    Return False
                Else

                    intYYYY = CType(("" & SplitTanggalAkhir(0)).ToString, Integer)
                    intMM = CType(("" & SplitTanggalAkhir(1)).ToString, Integer)
                    intDD = CType(("" & SplitTanggalAkhir(2)).ToString, Integer)

                    dateTanggalAkhir = New Date(intYYYY, intMM, intDD)

                End If

            End If
            'TanggalAkhir Selesai

            If dateTanggalAkhir < dateTanggalAwal Then
                PesanError = "Tanggal Akhir tidak boleh lebih kecil Tanggal Awal!"
                Return False
            End If

            'TimeSpan TanggalAwal dan TanggalAkhir tidak boleh lebih besar dari Limit
            Dim TSpan As TimeSpan = dateTanggalAkhir.Subtract(dateTanggalAwal)

            If TSpan.Days > MaxLimitPeriode Then
                PesanError = "Maksimal Periode Awal dan Akhir " & MaxLimitPeriode & " Hari!"
                Return False
            End If

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    Public Function LimitPeriodePartnerItemReport(ByVal TanggalAwal As String, ByVal TanggalAkhir As String, ByRef PesanError As String) As Boolean

        Try

            Dim LimitPeriodeTarikReport As String = ReadWebConfig("LimitPeriodePartnerItemReport")

            Dim MaxLimitPeriode As Integer = 90
            Try
                MaxLimitPeriode = CType(LimitPeriodeTarikReport, Integer)
            Catch ex As Exception
                MaxLimitPeriode = 90
            End Try

            Dim intYYYY As Integer = 0
            Dim intMM As Integer = 0
            Dim intDD As Integer = 0

            'TanggalAwal Mulai
            Dim dateTanggalAwal As Date = Nothing

            TanggalAwal = TanggalAwal.Trim

            Dim SplitTanggalJamAwal As String() = TanggalAwal.Split(" ") 'Jaga2 jika ada Jam nya
            If SplitTanggalJamAwal.Length < 1 Then
                PesanError = "Format Tanggal Awal yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                Return False
            Else

                Dim SplitTanggalAwal As String() = ("" & SplitTanggalJamAwal(0)).ToString.Trim.Split("-")

                If SplitTanggalAwal.Length < 3 Then
                    PesanError = "Format Tanggal Awal yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                    Return False
                Else

                    intYYYY = CType(("" & SplitTanggalAwal(0)).ToString, Integer)
                    intMM = CType(("" & SplitTanggalAwal(1)).ToString, Integer)
                    intDD = CType(("" & SplitTanggalAwal(2)).ToString, Integer)

                    dateTanggalAwal = New Date(intYYYY, intMM, intDD)

                End If

            End If
            'TanggalAwal Selesai

            'TanggalAkhir Mulai
            Dim dateTanggalAkhir As Date = Nothing

            TanggalAkhir = TanggalAkhir.Trim

            Dim SplitTanggalJamAkhir As String() = TanggalAkhir.Split(" ") 'Jaga2 jika ada Jam nya
            If SplitTanggalJamAkhir.Length < 1 Then
                PesanError = "Format Tanggal Akhir yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                Return False
            Else

                Dim SplitTanggalAkhir As String() = ("" & SplitTanggalJamAkhir(0)).ToString.Trim.Split("-")

                If SplitTanggalAkhir.Length < 3 Then
                    PesanError = "Format Tanggal Akhir yang valid harus YYYY-MM-DD (Contoh:2022-05-11)"
                    Return False
                Else

                    intYYYY = CType(("" & SplitTanggalAkhir(0)).ToString, Integer)
                    intMM = CType(("" & SplitTanggalAkhir(1)).ToString, Integer)
                    intDD = CType(("" & SplitTanggalAkhir(2)).ToString, Integer)

                    dateTanggalAkhir = New Date(intYYYY, intMM, intDD)

                End If

            End If
            'TanggalAkhir Selesai

            If dateTanggalAkhir < dateTanggalAwal Then
                PesanError = "Tanggal Akhir tidak boleh lebih kecil Tanggal Awal!"
                Return False
            End If

            'TimeSpan TanggalAwal dan TanggalAkhir tidak boleh lebih besar dari Limit
            Dim TSpan As TimeSpan = dateTanggalAkhir.Subtract(dateTanggalAwal)

            If TSpan.Days > MaxLimitPeriode Then
                PesanError = "Maksimal Periode Awal dan Akhir " & MaxLimitPeriode & " Hari!"
                Return False
            End If

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        End Try

    End Function

    Public Function GetBlastEmailTemplate(ByVal category As String, ByRef ErrMsg As String) As String

        Dim Template As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = category.ToString

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetBlastEmailTemplate(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                Template = respon(2).ToString

            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            Template = ""
        End If

        Return Template

    End Function

    Public Function GetBlastEmailCategory(ByRef ErrMsg As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.GetBlastEmailCategory(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

            Return dt

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally

            dt = Nothing
            GC.Collect()

        End Try

        Try

            If ErrMsg <> "" Then
                dt = New DataTable
                dt.TableName = "ERROR"
                dt.Columns.Add("RESPON")
                dt.Rows.Add(ErrMsg)
            End If

            Return dt

        Catch ex As Exception

            Return dt

        Finally

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function ThreePLOtherExpensesList() As DataTable

        Dim dt As DataTable = Nothing
        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)

        Dim MCon As MySqlConnection = Nothing

        Try

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            SqlQuery = "call sp_threeplotherexpenses_list();"
            SqlParam = New Dictionary(Of String, String)
            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            If Not MCon Is Nothing Then
                If MCon.State <> ConnectionState.Closed Then
                    MCon.Close()
                End If
                MCon.Dispose()
            End If

        End Try

        Return dt

    End Function

    Public Function ProsesRajaOngkirOverride(ByVal KecamatanList As String, ByVal Email As String, ByVal RJORequestId As String, ByVal UserId As String) As DataTable


        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(5)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = KecamatanList
            param(3) = Email
            param(4) = RJORequestId
            param(5) = UserId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.RajaOngkirProsesOverride(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"
            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

    Public Function GetRajaOngkirConstValue(ByVal Key As String) As DataTable


        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(2)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = Key

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.RajaOngkirGetConstValue(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"
            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

    Public Function GetRajaOngkirSubDistrictList(ByVal Type As String, ByVal RequestId As String) As DataTable

        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try

            Dim serv As New CoreService.Service
            Dim param() As Object
            Dim respon() As Object

            ReDim param(3)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS
            param(2) = Type
            param(3) = RequestId

            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")
            respon = serv.RajaOngkirGetSubDistrictList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)

            If respon(0) = "0" Then

                dt = ConvertStringToDatatable(respon(2))

                dt = dt.DefaultView.ToTable
                dt.TableName = "DATA"

            Else

                ErrMsg = respon(1)

            End If

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

    Public Sub SortDropDownAccountByCategorySubCategory(ByVal Type As String, ByVal FilterData As String(), ByVal ddl As DropDownList, ByVal chk As CheckBoxList, ByRef dtData As DataTable, ByRef dsData As DataSet, ByRef ErrorMessage As String)
        Try
            Dim dtTemp As New DataTable
            Dim dtEmpty As New DataTable

            Dim dtName As String = ""

            Select Case Type
                Case "CATEGORY" 'perubahan di ddlCategory mempengaruhi -> ddlSubCat dan ddlEcomm
                    For i As Integer = 0 To FilterData.Length - 1
                        dtName = "MAPPING_" & IIf(FilterData(i) = "", "0", FilterData(i).ToString)
                        If Not dsData.Tables.Contains(dtName) Then
                            dtEmpty = New DataTable
                            dtEmpty.TableName = dtName
                            dtEmpty.Columns.Add("ValueSubCategory", GetType(String))
                            dtEmpty.Columns.Add("DisplaySubCategory", GetType(String))
                            dsData.Tables.Add(dtEmpty)
                        End If
                        If i = 0 Then
                            dtTemp.Columns.Add("ValueSubCategory", GetType(String))
                            dtTemp.Columns.Add("DisplaySubCategory", GetType(String))
                            dtTemp.Rows.Add("", "")
                        End If
                        dtTemp.Merge(dsData.Tables(dtName))
                    Next

                    ddlBound(ddl, dtTemp, "DisplaySubCategory", "ValueSubCategory")

                    dtTemp.DefaultView.RowFilter = "ValueSubCategory <> ''"
                    dtTemp = dtTemp.DefaultView.ToTable
                    chkBound(chk, dtTemp, "DisplaySubCategory", "ValueSubCategory")


                Case "SUBCATEGORY-ACCOUNT" 'Perubahan di ddlSubCat mempengaruhi -> ddlEcomm

                    For i As Integer = 0 To FilterData.Length - 1
                        dtName = "SUBCAT_" & IIf(FilterData(i) = "", "0", FilterData(i).ToString)
                        If Not dsData.Tables.Contains(dtName) Then
                            dtEmpty = New DataTable
                            dtEmpty.TableName = dtName
                            dtEmpty.Columns.Add("ValueAccount", GetType(String))
                            dtEmpty.Columns.Add("DisplayAccount", GetType(String))
                            dsData.Tables.Add(dtEmpty)
                        End If
                        If i = 0 Then
                            dtTemp.Columns.Add("ValueAccount", GetType(String))
                            dtTemp.Columns.Add("DisplayAccount", GetType(String))
                            dtTemp.Rows.Add("", "")
                        End If
                        dtTemp.Merge(dsData.Tables(dtName))
                    Next


                    ddlBound(ddl, dtTemp, "DisplayAccount", "ValueAccount")

                    dtTemp.DefaultView.RowFilter = "ValueAccount <> ''"
                    dtTemp = dtTemp.DefaultView.ToTable
                    chkBound(chk, dtTemp, "DisplayAccount", "ValueAccount")


                Case "CATEGORY-ACCOUNT" 'perubahan di ddlCategory mempengaruhi -> ddlEcomm
                    For i As Integer = 0 To FilterData.Length - 1
                        dtName = "CAT_" & IIf(FilterData(i) = "", "0", FilterData(i).ToString)
                        If Not dsData.Tables.Contains(dtName) Then
                            dtEmpty = New DataTable
                            dtEmpty.TableName = dtName
                            dtEmpty.Columns.Add("ValueAccount", GetType(String))
                            dtEmpty.Columns.Add("DisplayAccount", GetType(String))
                            dsData.Tables.Add(dtEmpty)
                        End If
                        If i = 0 Then
                            dtData.Columns.Add("ValueAccount", GetType(String))
                            dtData.Columns.Add("DisplayAccount", GetType(String))
                            dtData.Rows.Add("", "")
                        End If
                        dtData.Merge(dsData.Tables(dtName))
                    Next

                    ddlBound(ddl, dtData, "DisplayAccount", "ValueAccount")

                    dtData.DefaultView.RowFilter = "ValueAccount <> ''"
                    dtTemp = dtData.DefaultView.ToTable
                    chkBound(chk, dtTemp, "DisplayAccount", "ValueAccount")


                Case Else
                    ErrorMessage = "Tipe proses " & Type & " tidak ditemukan!"

            End Select

        Catch ex As Exception
            ErrorMessage = ex.Message
        End Try

    End Sub

    Public Function GetMstTrackingStatusList(ByVal Criteria As String, ByRef ErrorMessage As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call sp_msttrackingstatus_ddl(@Criteria)"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Criteria", Criteria.ToUpper)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam, ErrorMessage)

            Return dt

        Catch ex As Exception

            ErrorMessage = ex.Message
            Return dt

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            SqlParam = Nothing

            GC.Collect()

        End Try

    End Function

    Public Function GetPinRequestReasonList() As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try
            SqlQuery = "call sp_pinrequest_reason_list();"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)
            dt.TableName = "DATA"

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

    Public Function GetPinRequestDescriptionList() As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Dim ErrMsg As String = ""

        Try
            SqlQuery = "call sp_pinrequest_description_list();"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)
            dt.TableName = "DATA"

        Catch ex As Exception

            ErrMsg = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        If ErrMsg <> "" Then
            dt = New DataTable
            dt.TableName = "ERROR"
            dt.Columns.Add("RESPON")
            dt.Rows.Add(ErrMsg)
        End If

        Return dt

    End Function

#Region "Incentive Driver"

    Public Function IncentiveDriverGetCabangReportList(ByVal Code As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_cabang_report_list`(@Code);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Code", Code)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetCabangList(ByVal Cabang As String, ByVal Hub As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_cabang_list`(@Cabang, @Hub);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)
            SqlParam.Add("@Hub", Hub)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetJenisKirimanList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_deliverytype_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetRitaseList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_ritase_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetJenisArmadaList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_vehicletype_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetNomorKendaraanList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_nopol_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetDriverList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_driver_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetDelimanList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_deliman_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetMemberList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_member_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDisableInputList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_disableinput_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveValidateLoadingMalam(ByVal Cabang As String, ByVal Tanggal As String, ByVal NoPol As String) As String

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing
        Dim Result As String = "ERROR"

        Try

            SqlQuery = "call `sp_incdrv_validate_loadingmalam`(@Cabang, @Tanggal, @NoPol);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)
            SqlParam.Add("@Tanggal", Tanggal)
            SqlParam.Add("@NoPol", NoPol)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                Result = ""
                If dt.Rows.Count > 0 Then
                    Result = dt.Rows(0).Item("IncDrvId").ToString
                End If
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Result = "ERROR"

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

        Return Result

    End Function

    Public Function IncentiveValidateDoubleInput(ByVal Cabang As String, ByVal Tanggal As String, ByVal Ritase As String _
    , ByVal NoPol As String, ByVal DriverId As String, ByVal DelimanId As String) As String

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing
        Dim Result As String = "ERROR"

        Try

            SqlQuery = "call `sp_incdrv_validate_doubleinput`(@Cabang, @Tanggal, @Ritase, @NoPol, @DriverId, @DelimanId);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)
            SqlParam.Add("@Tanggal", Tanggal)
            SqlParam.Add("@Ritase", Ritase)
            SqlParam.Add("@NoPol", NoPol)
            SqlParam.Add("@DriverId", DriverId)
            SqlParam.Add("@DelimanId", DelimanId)

            MCon = ObjCon.SetConn_Master
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                Result = ""
                If dt.Rows.Count > 0 Then
                    Result = dt.Rows(0).Item("IncDrvId").ToString
                End If
            End If

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Result = "ERROR"

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

        Return Result

    End Function

    Public Function IncentiveDriverReviseReasonList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_revisereason_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function IncentiveDriverGetOtherValidationList(ByVal Cabang As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try

            SqlQuery = "call `sp_incdrv_othervalidation_list`(@Cabang);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Cabang", Cabang)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

#End Region

#Region "Resi Konsol"

    Public Function NoResiKonsolGetTemplateList(Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try
            SqlQuery = "call `noresikonsol_template_list`();"

            SqlParam = New Dictionary(Of String, String)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
                dt = New DataTable
                dt.TableName = "ERROR"
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function NoResiKonsolGetRuteKonsList(ByVal Code As String) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try
            SqlQuery = "call `noresikonsol_rutekons_list`(@Code);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Code", Code)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
                dt = New DataTable
                dt.TableName = "ERROR"
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function NoResiKonsolGetExpeditionList(ByVal Template As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try
            SqlQuery = "call `noresikonsol_expedition_list`(@Template);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Template", Template)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
                dt = New DataTable
                dt.TableName = "ERROR"
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

    Public Function NoResiKonsolGetKendaraanList(ByVal Template As String, Optional ByVal ActiveOnly As Boolean = True) As DataTable

        Dim SqlQuery As String = ""
        Dim SqlParam As New Dictionary(Of String, String)
        Dim MCon As MySqlConnection = Nothing
        Dim dt As DataTable = Nothing

        Try
            SqlQuery = "call `noresikonsol_kendaraan_list`(@Template);"

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Template", Template)

            MCon = ObjCon.SetConn_Slave1
            If MCon.State <> ConnectionState.Open Then
                MCon.Open()
            End If

            dt = ObjSQL.SQLInsertIntoDatatable(MCon, SqlQuery, SqlParam)

            If dt Is Nothing Then
                dt = New DataTable
                dt.TableName = "ERROR"
            Else
                If ActiveOnly Then
                    dt.DefaultView.RowFilter = "IsActive = '1'"
                    dt = dt.DefaultView.ToTable
                End If
            End If

            Return dt

        Catch ex As Exception

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        Finally

            Try
                MCon.Close()
            Catch ex As Exception
            End Try

            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

            dt = Nothing
            GC.Collect()

        End Try

    End Function

#End Region

    Public Function StringDayOfWeekIna(ByVal StringOri As String) As String

        Dim Result As String = ""

        Select Case StringOri.ToUpper
            Case "SUNDAY"
                Result = "Minggu"
            Case "MONDAY"
                Result = "Senin"
            Case "TUESDAY"
                Result = "Selasa"
            Case "WEDNESDAY"
                Result = "Rabu"
            Case "THURSDAY"
                Result = "Kamis"
            Case "FRIDAY"
                Result = "Jumat"
            Case "SATURDAY"
                Result = "Sabtu"
            Case Else
                Result = "XXX"
        End Select

        Return Result

    End Function

    Public Function GetDeliveryDateList() As DataTable

        Dim dt As New DataTable
        dt.Columns.Add("Value")
        dt.Columns.Add("Display")

        dt.Rows.Add()
        dt.Rows(dt.Rows.Count - 1).Item("Value") = ""
        dt.Rows(dt.Rows.Count - 1).Item("Display") = ""

        For i As Integer = 1 To 14
            Dim DateValue As Date = DateAdd(DateInterval.Day, i - 1, Date.Now)
            Dim ValueDate As String = Format(DateValue, "yyyy-MM-dd")
            Dim DisplayDate As String = StringDayOfWeekIna(DateValue.DayOfWeek.ToString) & ", " & Format(DateValue, "dd MMM yy")

            dt.Rows.Add()
            dt.Rows(dt.Rows.Count - 1).Item("Value") = ValueDate
            dt.Rows(dt.Rows.Count - 1).Item("Display") = DisplayDate
        Next

        Return dt

    End Function

    Public Function GetDeliveryTimeList() As DataTable

        Dim dt As New DataTable
        dt.Columns.Add("Value")
        dt.Columns.Add("Display")

        dt.Rows.Add()
        dt.Rows(dt.Rows.Count - 1).Item("Value") = ""
        dt.Rows(dt.Rows.Count - 1).Item("Display") = ""

        For i As Integer = 0 To 23
            dt.Rows.Add()
            dt.Rows(dt.Rows.Count - 1).Item("Value") = i.ToString.PadLeft(2, "0")
            dt.Rows(dt.Rows.Count - 1).Item("Display") = i.ToString.PadLeft(2, "0")
        Next

        Return dt

    End Function

    Public Function SendHTTP(ByVal IDTerminal As String, ByVal TrxToko As String,
                   ByVal tipeReg As String,
                   ByVal URL As String, ByVal Parameter As String, ByVal userAgent As String,
                   ByVal TipeEncode As System.Text.Encoding, ByVal TimeOutWeb As String, ByVal JawabanDebug As String,
                   Optional ByVal content As String = "application/x-www-form-urlencoded", Optional ByVal isTls As Boolean = False,
                   Optional ByVal CustomHeaders As String() = Nothing, Optional ByVal CustomMethod As String = "POST",
                   Optional ByVal ExpectedSuccessCode As HttpStatusCode = HttpStatusCode.OK) As String

        Dim hasil As String = ""

        Dim myRequest As HttpWebRequest
        Dim response As HttpWebResponse
        Dim byteArray As Byte()
        Dim dataStream As Stream
        Dim reader As StreamReader

        Try
            If JawabanDebug.Trim = "" Then
                myRequest = HttpWebRequest.Create(URL)

                If isTls = True Then
                    'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or CType(768, SecurityProtocolType) Or CType(3072, SecurityProtocolType)
                    ServicePointManager.Expect100Continue = False
                End If

                If URL.StartsWith("https") Then
                    'ulangi:
                    ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CertificateValidationCallBack)
                    Dim networkCredential As New NetworkCredential("INDOMARET", "#2tD$75/U*8z3B9+4CZ6")
                    myRequest.Credentials = networkCredential
                End If

                myRequest.Timeout = TimeOutWeb * 1000
                myRequest.Method = "POST"
                If CustomMethod <> "POST" Then
                    myRequest.Method = CustomMethod
                End If

                Dim encoding As System.Text.Encoding
                encoding = TipeEncode
                byteArray = encoding.GetBytes(Parameter)
                If myRequest.Method <> "GET" Then
                    myRequest.ContentType = content '"application/x-www-form-urlencoded"
                End If
                myRequest.ContentLength = byteArray.Length

                Try
                    If Not CustomHeaders Is Nothing Then
                        For i As Integer = 0 To CustomHeaders.Length - 1
                            Dim CustomHeader As String() = CustomHeaders(i).Split("|")
                            myRequest.Headers.Add(CustomHeader(0), CustomHeader(1))
                        Next
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                End Try

                If userAgent <> "" Then
                    myRequest.UserAgent = userAgent
                End If

                If myRequest.Method <> "GET" Then
                    dataStream = myRequest.GetRequestStream()
                    dataStream.Write(byteArray, 0, byteArray.Length)
                    dataStream.Close()
                End If

                response = myRequest.GetResponse

                'If response.StatusCode = HttpStatusCode.OK Then
                If response.StatusCode = ExpectedSuccessCode Then

                    reader = New StreamReader(response.GetResponseStream())

                    'hasil = reader.ReadToEnd
                    Using (reader)
                        hasil = reader.ReadToEnd
                    End Using

                End If

            End If

        Catch ex As WebException
            Dim statusCode As HttpStatusCode
            Dim StatusDesc As String
            If (ex.Status = WebExceptionStatus.ProtocolError) Then
                Dim Errresponse As WebResponse = ex.Response
                Using (Errresponse)
                    Dim httpResponse As HttpWebResponse = CType(Errresponse, HttpWebResponse)
                    statusCode = httpResponse.StatusCode
                    StatusDesc = httpResponse.StatusDescription
                    Try
                        reader = New StreamReader(Errresponse.GetResponseStream())
                        Using (reader)
                            hasil = reader.ReadToEnd '& "Status Description = " & httpResponse.StatusDescription ' HttpWebResponse.StatusDescription
                        End Using
                    Catch err As Exception
                    End Try
                End Using
            End If

        Catch exx As Exception
            hasil = ""

        Finally
        End Try

        Return hasil.Trim

    End Function

    Public Function GoogleGeocodingLocation(ByVal dtRequest As DataTable, ByRef ErrorMessage As String) As DataTable

        Dim Method As String = "GoogleGeocodingLocation"

        Dim dtResult As DataTable = Nothing

        Try

            If dtRequest.Rows.Count > 0 Then

                Dim MustHaveColumns As Integer = 0
                Try
                    For c As Integer = 0 To dtRequest.Columns.Count - 1
                        If dtRequest.Columns(c).ColumnName.ToUpper = "REQADDRESS" _
                        Or dtRequest.Columns(c).ColumnName.ToUpper = "REQPOSTALCODE" _
                        Or dtRequest.Columns(c).ColumnName.ToUpper = "REQLATITUDE" _
                        Or dtRequest.Columns(c).ColumnName.ToUpper = "REQLONGITUDE" Then
                            MustHaveColumns += 1
                        End If
                    Next
                Catch ex As Exception

                End Try

                If MustHaveColumns < 4 Then
                    ErrorMessage = "Kolom dtRequest Tidak Lengkap, harus ada kolom REQADDRESS, REQPOSTALCODE, REQLATITUDE, REQLONGITUDE"
                Else

                    dtResult = dtRequest.Copy

                    dtResult.Columns.Add("Status")

                    Dim CompleteUrl As String = ""
                    Dim Url As String = "https://maps.googleapis.com/maps/api/geocode/json"
                    'Dim Key As String = "AIzaSyBbgx8BaEhRtF7558sdDsdS077Iizc6ZV0"
                    Dim Key As String = ("" & ConfigurationManager.AppSettings("GoogleGeocodingKey")).Trim
                    Key = "key=" & Key

                    Dim Response As String = ""
                    Dim JSON_Res_Geocoding As New JSON_GGL_ResponseGeocoding

                    Dim ReqAddress As String = ""
                    Dim ReqPostalCode As String = ""
                    Dim ReqLatitude As String = ""
                    Dim ReqLongitude As String = ""

                    For i As Integer = 0 To dtResult.Rows.Count - 1

                        ReqAddress = ("" & dtResult.Rows(i).Item("REQADDRESS")).ToString
                        ReqPostalCode = ("" & dtResult.Rows(i).Item("REQPOSTALCODE")).ToString
                        ReqLatitude = ("" & dtResult.Rows(i).Item("REQLATITUDE")).ToString
                        ReqLongitude = ("" & dtResult.Rows(i).Item("REQLONGITUDE")).ToString

                        If ReqAddress = "" Then
                            dtResult.Rows(i).Item("Status") = "Alamat Kosong!"
                            ErrorMessage = "Alamat Kosong!"
                        Else

                            If ReqLatitude = "" Or ReqLongitude = "" Then

                                If Not ReqAddress.Contains(ReqPostalCode) Then
                                    If ReqPostalCode.Trim <> "" Then
                                        ReqAddress = ReqAddress & ", " & ReqPostalCode
                                    End If
                                End If
                                If Not ReqAddress.ToUpper.Contains("INDONESIA") Then
                                    ReqAddress = ReqAddress & ", INDONESIA"
                                End If

                                Dim Address As String = "address=" & ReqAddress

                                'cek dulu di table cache
                                Dim dtCache As DataTable = CacheCoordinateGet(ReqAddress)
                                If dtCache.Rows.Count > 0 Then

                                    dtResult.Rows(i).Item("Status") = "OK"
                                    dtResult.Rows(i).Item("ReqLatitude") = dtCache.Rows(0).Item("Latitude").ToString
                                    dtResult.Rows(i).Item("ReqLongitude") = dtCache.Rows(0).Item("Longitude").ToString

                                Else

                                    CompleteUrl = ""

                                    CompleteUrl &= Url
                                    CompleteUrl &= "?" & Key
                                    CompleteUrl &= "&" & Address

                                    Response = SendHTTP("", "", "", CompleteUrl, "", "", System.Text.Encoding.UTF8, "60", "", , IsTlsGoogleApi)

                                    JSON_Res_Geocoding = New JSON_GGL_ResponseGeocoding
                                    JSON_Res_Geocoding = Newtonsoft.Json.JsonConvert.DeserializeObject(Of JSON_GGL_ResponseGeocoding)(Response)

                                    If JSON_Res_Geocoding.status.ToUpper = "OK" Then

                                        If JSON_Res_Geocoding.results.Length > 0 Then
                                            dtResult.Rows(i).Item("Status") = JSON_Res_Geocoding.status
                                            dtResult.Rows(i).Item("ReqLatitude") = ("" & JSON_Res_Geocoding.results(0).geometry.location.lat).ToString
                                            dtResult.Rows(i).Item("ReqLongitude") = ("" & JSON_Res_Geocoding.results(0).geometry.location.lng).ToString

                                            'simpan di table cache
                                            CacheCoordinateSet(ReqAddress, ("" & JSON_Res_Geocoding.results(0).geometry.location.lat).ToString, ("" & JSON_Res_Geocoding.results(0).geometry.location.lng).ToString)

                                        Else

                                            Try
                                                dtResult.Rows(i).Item("Status") = "Koordinat Alamat Tidak Ditemukan. Cek Kembali Alamat"
                                            Catch ex As Exception
                                            End Try

                                        End If

                                    Else
                                        dtResult.Rows(i).Item("Status") = JSON_Res_Geocoding.status

                                        Try
                                            If dtResult.Rows(i).Item("Status").ToString.ToUpper.Contains("ZERO") _
                                            And dtResult.Rows(i).Item("Status").ToString.ToUpper.Contains("RESULT") Then
                                                dtResult.Rows(i).Item("Status") = "Koordinat Alamat Tidak Bisa Ditentukan. Cek Kembali Alamat"
                                            End If

                                            If dtResult.Rows(i).Item("Status").ToString.ToUpper.Contains("NOT") _
                                            And dtResult.Rows(i).Item("Status").ToString.ToUpper.Contains("FOUND") Then
                                                dtResult.Rows(i).Item("Status") = "Koordinat Alamat Tidak Bisa Ditentukan. Cek Kembali Alamat"
                                            End If
                                        Catch ex As Exception
                                        End Try

                                    End If

                                End If 'dari If dtCache.Rows.Count > 0

                            End If 'dari If ReqLatitude = "" Or ReqLongitude = ""

                        End If 'dari If ReqAddress = ""

                    Next

                End If 'dari If MustHaveColumns

            End If

        Catch ex As Exception

            ErrorMessage = ex.Message

        End Try

        Return dtResult

    End Function

    Private Function CacheCoordinateGet(ByVal Address As String) As DataTable

        Dim Result As New DataTable

        Dim MCon As MySqlConnection

        Try
            MCon = MasterMCon.Clone
            MCon.Open()

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@address", Address)

            Result = ObjSQL.ExecDatatableWithParam(MCon, "Select * From GeoCodingLatLon Where Address = @address And UpdTime >= date_add(now(), interval -3 month)", SqlParam)

        Catch ex As Exception
            Result = New DataTable

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    Private Sub CacheCoordinateSet(ByVal Address As String, ByVal Latitude As String, ByVal Longitude As String)

        Dim MCon As MySqlConnection

        Try
            MCon = MasterMCon.Clone
            MCon.Open()

            Dim sb As New StringBuilder
            sb.Append(" Insert Into GeoCodingLatLon (")
            sb.Append(" Address, Latitude, Longitude")
            sb.Append(" , UpdTime, UpdUser")
            sb.Append(" ) values (")
            sb.Append(" @address, @lat, @lon")
            sb.Append(" , now(), 'system'")
            sb.Append(" ) On Duplicate Key Update UpdTime = now()")

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@address", Address)
            SqlParam.Add("@lat", Latitude)
            SqlParam.Add("@lon", Longitude)

            Dim ObjSQL As New ClsSQL
            ObjSQL.ExecNonQueryWithParam(MCon, sb.ToString, SqlParam)

        Catch ex As Exception

        Finally
            Try
                MCon.Close()
            Catch ex As Exception
            End Try
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

    End Sub

    Public Function GetHubExpeditionList(ByVal ViewStateCurrentWebCode As String, ByVal Hub As String, ByVal ProcessType As String, ByVal ShowAddInfo As String, ByVal ReferenceNo As String, ByRef ErrorMessage As String) As DataTable

        'If Hub = "" Then
        'ErrorMessage = "Session Nothing!"
        'Return Nothing
        'End If

        'If ViewStateCurrentWebCode <> Hub Then
        'ErrorMessage = "Login Awal : " & ViewStateCurrentWebCode & ", Login Sekarang : " & Hub
        'Return Nothing
        'End If

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object = Nothing

            'ProcessType:
            'SERAH
            'TERIMA
            'JEMPUT REKANAN
            'JEMPUT EKSPEDISI
            'RESI KONSOL

            ReDim param(5)
            'ReDim param(1)
            param(0) = UserWS
            param(1) = PassWS
            param(2) = Hub
            param(3) = ProcessType
            param(4) = ShowAddInfo
            param(5) = ReferenceNo

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

            Dim i As Integer = 1
            Dim success As Boolean = False
            While i <= maxTryWS And success = False

                Try

                    respon = serv.GetHubExpeditionList(AppName, AppVersion, param)
                    success = True

                Catch ex As Exception

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                    i = i + 1

                End Try

            End While

            If success Then

                If respon(0) = "0" Then

                    Dim dt As DataTable = ConvertStringToDatatable(respon(2))
                    Return dt

                Else

                    ErrorMessage = respon(1)
                    Return Nothing

                End If

            Else

                ErrorMessage = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                WriteTracelogTxt(Pesan)

                Return Nothing

            End If

        Catch ex As Exception

            ErrorMessage = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return Nothing

        End Try

    End Function

    Public Function PrintAWB3PL(ByVal Header() As String, ByVal Detail As DataTable, ByVal NoAWB3PL As String, ByVal ViewStateCurrentWebCode As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentWebCode = "" Then
            PesanError = "Session Nothing!"
            Return False
        End If

        If ViewStateCurrentWebCode <> CurrentWebCode Then
            PesanError = "Login Awal : " & ViewStateCurrentWebCode & ", Login Sekarang : " & CurrentWebCode
            Return False
        End If

        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Try

            Dim NamaRpt As String = "AWB3PL.rpt"

            Dim TanggalProses As String = Header(0)
            Dim DocumentName As String = Header(1)
            Dim DocumentNo As String = Header(2)
            Dim PetugasHUBID As String = Header(3)
            Dim PetugasHUBName As String = Header(4)
            Dim TPLName As String = Header(5)
            Dim AWB3PL As String = Header(6)
            Dim TglAWB3PL As String = Header(7)
            Dim Asal As String = Header(8)
            Dim Tujuan As String = Header(9)
            Dim Berat As String = Header(10)
            Dim LblBerat As String = Header(11)

            Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

            rpt.Load(FileRptPath)
            rpt.SummaryInfo.ReportTitle = "AWB3PL"
            rpt.PrintOptions.PaperSize = PaperSize.PaperA4
            rpt.SetDataSource(Detail)

            Dim crParameterFieldDefinitions As ParameterFieldDefinitions
            Dim crParameterFieldLocation As ParameterFieldDefinition
            Dim crParameterDiscreteValue As ParameterDiscreteValue
            Dim crParameterValues As ParameterValues

            crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

            crParameterFieldLocation = crParameterFieldDefinitions.Item("TanggalProses")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = TanggalProses
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = DocumentName
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentNo")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = DocumentNo
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasHUBID")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = PetugasHUBID
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("PetugasHUBName")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = PetugasHUBName
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("3PLName")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = TPLName
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("AWB3PL")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = AWB3PL
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("TglAWB3PL")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = TglAWB3PL
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("Asal")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = Asal
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("Tujuan")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = Tujuan
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("Berat")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = Berat
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            crParameterFieldLocation = crParameterFieldDefinitions.Item("LblBerat")
            crParameterValues = crParameterFieldLocation.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = LblBerat
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

            Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
            Dim FullPathReportFile As String = PathReportFile & "\BuktiProsesAWB3PL_" & CurrentWebCode & "_AWB3PL_" & NoAWB3PL & "_stamp_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".pdf"
            FullPathReportFile = FullPathReportFile.Replace("\\", "\")

            Try
                Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                Dim DownloadGeneratedFile As String = ReadWebConfig("DownloadGeneratedFile", False)
                If PrintToPDF = "1" Or DownloadGeneratedFile = "1" Then
                    If Not Directory.Exists(PathReportFile) Then
                        Directory.CreateDirectory(PathReportFile)
                    End If

                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                    CreatedFilePath = FullPathReportFile
                End If
            Catch ex As Exception

                PesanError = ex.Message

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String = ""
                Pesan = "Dari ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                WriteTracelogTxt(Pesan)

            End Try

            Try

                Dim PrintToPrinter As String = ReadWebConfig("PrintToPrinter", False)

                If PrintToPrinter = "1" Then
                    Dim HubPrinterName As String = SetHubPrinterName()
                    rpt.PrintOptions.PrinterName = HubPrinterName
                    rpt.PrintToPrinter(1, True, 0, 0)
                End If

            Catch ex As Exception

                Dim prd As New System.Drawing.Printing.PrintDocument
                rpt.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName
                rpt.PrintToPrinter(1, True, 0, 0)

            End Try

            Try
                crParameterFieldDefinitions.Dispose()
                crParameterFieldLocation.Dispose()
            Catch
            End Try

            Return True

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False

        Finally
            rpt.Close()
            rpt.Dispose()
        End Try

        Return False

    End Function

    Public Function PrintSerahPackingList(ByVal AWB3PL As String, ByVal TipeCetak As String, ByVal ViewStateCurrentWebCode As String, ByVal CurrentWebCode As String, ByVal CurrentWebName As String, ByVal Expedition As String, ByVal ExpeditionName As String, ByRef CreatedFilePath As String, ByRef PesanError As String) As Boolean

        If CurrentWebCode = "" Then
            PesanError = "Session Nothing!"
            Return False
        End If

        If ViewStateCurrentWebCode <> CurrentWebCode Then
            PesanError = "Login Awal : " & ViewStateCurrentWebCode & ", Login Sekarang : " & CurrentWebCode
            Return False
        End If

        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Try

            Dim serv As New CoreService.Service
            serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

            Dim respon() As Object = Nothing

            Dim param(4) As Object
            param(0) = UserWS
            param(1) = PassWS
            param(2) = CurrentWebCode
            param(3) = Expedition
            param(4) = AWB3PL

            Dim i As Integer = 1
            Dim success As Boolean = False
            While i <= maxTryWS And success = False

                Try

                    respon = serv.PrintSerahPackingList(AppName, AppVersion, param)
                    success = True

                Catch ex As Exception

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                    i = i + 1

                End Try

            End While

            If success Then

                If respon(0) = "0" Then

                    Dim NamaRpt As String = "PackingList.rpt"
                    Dim dt As New DataTable("PackingList")

                    Dim strReprint As String = ""
                    If TipeCetak.ToUpper = "REPRINT" Then
                        strReprint = "REPRINT "
                    End If

                    Dim DocumentName As String = strReprint & "PACKING LIST"
                    Dim NamaEkspedisi As String = "" & ExpeditionName
                    Dim HubName As String = "" & CurrentWebName

                    dt = ConvertStringToDatatableWithBarcode(respon(2))

                    'Perlu kolom QtyKoli_Int dan WeightToSum_Dbl
                    'Karena QtyKoli dan WeightToSum DataType String, saat di SUM di RPT hasil nya salah
                    If Not IsNothing(dt) Then

                        dt.Columns.Add("QtyKoli_Int", GetType(Integer))
                        dt.Columns.Add("WeightToSum_Dbl", GetType(Double))

                        For j As Integer = 0 To dt.Rows.Count - 1
                            dt.Rows(j).Item("QtyKoli_Int") = dt.Rows(j).Item("QtyKoli")
                            dt.Rows(j).Item("WeightToSum_Dbl") = dt.Rows(j).Item("WeightToSum")
                        Next

                    End If

                    Dim FileRptPath As String = ReadWebConfig("PathRpt", True) & NamaRpt

                    rpt.Load(FileRptPath)
                    rpt.SummaryInfo.ReportTitle = "Packing List"
                    rpt.PrintOptions.PaperSize = PaperSize.PaperA4
                    rpt.SetDataSource(dt)

                    Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                    Dim crParameterFieldLocation As ParameterFieldDefinition
                    Dim crParameterDiscreteValue As ParameterDiscreteValue
                    Dim crParameterValues As ParameterValues

                    crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields

                    crParameterFieldLocation = crParameterFieldDefinitions.Item("DocumentName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = DocumentName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

                    crParameterFieldLocation = crParameterFieldDefinitions.Item("NamaEkspedisi")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = NamaEkspedisi
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

                    crParameterFieldLocation = crParameterFieldDefinitions.Item("HubName")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = HubName
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

                    crParameterFieldLocation = crParameterFieldDefinitions.Item("AWB3PL")
                    crParameterValues = crParameterFieldLocation.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = AWB3PL
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldLocation.ApplyCurrentValues(crParameterValues)

                    Dim PathReportFile As String = ReadWebConfig("PathReportFile", False)
                    Dim FullPathReportFile As String = PathReportFile & "\PackingList_" & CurrentWebCode & "_AWB3PL_" & AWB3PL & "_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".pdf"
                    FullPathReportFile = FullPathReportFile.Replace("\\", "\")

                    Try
                        Dim PrintToPDF As String = ReadWebConfig("PrintToPDF", False)
                        Dim DownloadGeneratedFile As String = ReadWebConfig("DownloadGeneratedFile", False)
                        If PrintToPDF = "1" Or DownloadGeneratedFile = "1" Then
                            If Not Directory.Exists(PathReportFile) Then
                                Directory.CreateDirectory(PathReportFile)
                            End If

                            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, FullPathReportFile)
                            CreatedFilePath = FullPathReportFile
                        End If
                    Catch ex As Exception

                        PesanError = ex.Message

                        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                        Dim Pesan As String = ""
                        Pesan = "Dari Halaman ClsFungsi, Proses " & methodName & " Export to PDF, Error : " & ex.Message
                        WriteTracelogTxt(Pesan)

                    End Try

                    Try

                        Dim PrintToPrinter As String = ReadWebConfig("PrintToPrinter", False)

                        If PrintToPrinter = "1" Then
                            Dim HubPrinterName As String = SetHubPrinterName()
                            rpt.PrintOptions.PrinterName = HubPrinterName
                            rpt.PrintToPrinter(1, True, 0, 0)
                        End If

                    Catch

                        Dim prd As New System.Drawing.Printing.PrintDocument
                        rpt.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName
                        rpt.PrintToPrinter(1, True, 0, 0)

                    End Try

                    Try
                        crParameterFieldDefinitions.Dispose()
                        crParameterFieldLocation.Dispose()
                    Catch
                    End Try

                    Return True
                Else
                    PesanError = respon(1)

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & PesanError
                    WriteTracelogTxt(Pesan)

                    Return False
                End If

            Else

                PesanError = "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"

                Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                Dim Pesan As String
                Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal Konek ke Coreservice " & maxTryWS.ToString & "x"
                WriteTracelogTxt(Pesan)

                Return False

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

            Return False
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try

        Return False

    End Function

    Public Function SetHubPrinterName() As String

        Dim Result As String = ""

        Try
            Result = ReadCustomConfig("CustomConfigPrinterPath", "PRINTERNAME")
        Catch ex As Exception
            Result = ReadWebConfig("PrinterName", False)
        End Try

        Return Result

    End Function

    Public Function ReadCustomConfig(ByVal ConfigPath As String, ByVal NamaSetting As String) As String

        Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name

        Try

            Dim CustomConfigPath As String = ReadWebConfig(ConfigPath)

            Dim fileReader As String = ""
            Try
                fileReader = My.Computer.FileSystem.ReadAllText(CustomConfigPath)
            Catch ex As Exception
                fileReader = ""
            End Try

            If fileReader <> "" Then

                Dim FileData() = fileReader.Replace(vbCrLf, vbLf).Split(vbLf)

                For Each splited As String In FileData
                    Try

                        Dim Config As String = splited.Split("=")(0)
                        Dim Value As String = splited.Split("=")(1)

                        If Config.ToUpper = NamaSetting.ToUpper Then
                            Return Value
                        End If
                    Catch ex As Exception
                    End Try
                Next

            Else
                WriteTracelogTxt("Dari ClsFungsi, Proses " & methodName & ", Error : " & "Gagal baca file config")
            End If

            Return ""

        Catch ex As Exception

            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)
            Return ""

        End Try

    End Function

    Public Function GetHubList2(ByRef PesanError As String) As DataTable

        Dim dt As DataTable = Nothing

        Try

            Dim serv As New LocalCore
            Dim param() As Object
            Dim respon() As Object = Nothing

            ReDim param(1)
            param(0) = ClsWebVer.UserWS
            param(1) = ClsWebVer.PassWS

            'serv.Url = System.Configuration.ConfigurationManager.AppSettings("CoreServiceURL")

            Dim i As Integer = 1
            Dim success As Boolean = False
            While i <= maxTryWS And success = False

                Try

                    respon = serv.GetHubList(ClsWebVer.AppName, ClsWebVer.AppVersion, param)
                    success = True

                Catch ex As Exception

                    Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
                    Dim Pesan As String = ""
                    Pesan = "Dari ClsFungsi, Proses " & methodName & ", Percobaan Konek Coreservice ke " & i.ToString & ", Error : " & ex.Message
                    WriteTracelogTxt(Pesan)

                    i = i + 1

                End Try

            End While

            If success Then

                If respon(0) = "0" Then

                    dt = ConvertStringToDatatable(respon(2))

                Else

                    PesanError = respon(1)

                End If

            End If

        Catch ex As Exception

            PesanError = ex.Message

            Dim methodName As String = System.Reflection.MethodBase.GetCurrentMethod.Name
            Dim Pesan As String = ""
            Pesan = "Dari ClsFungsi, Proses " & methodName & ", Error : " & ex.Message
            WriteTracelogTxt(Pesan)

        End Try

        Return dt

    End Function


    Public Function GetPackageList(ByVal Parameter As String, ByVal Jenis As String) As String
        'jenis = XML / JSON
        Jenis = Jenis.ToUpper

        Dim serv As New LocalCore
        Dim servParam(3) As Object

        Dim balikan() As String                'UNTUK BALIKAN DARI CORE SERVICE ( DALAM BENTUK STRING ARRAY )
        Dim sBalikan As String = ""             'UNTUK BALIKAN DENGAN FORMAT XML ATAU JSON

        Dim AppName As String = ""
        Dim AppVersion As String = ""

        Dim ResponseOK As Boolean = False

        Try

            Dim url As String = ConfigurationManager.AppSettings("CoreServiceURL")

            'serv.Url = url
            'serv.Timeout = ConfigurationManager.AppSettings("timeout")

            If Jenis = "JSON" Then

                Dim _ObjRequest As New _clsPackageList

                _ObjRequest = JsonConvert.DeserializeObject(Of _clsPackageList)(Parameter)

                If _ObjRequest Is Nothing Then
                    'parameter tidak valid
                    servParam = Nothing

                Else
                    AppName = ("" & _ObjRequest.AppName)
                    AppVersion = ("" & _ObjRequest.AppVersion)

                    servParam(0) = ("" & _ObjRequest.User)
                    If ("" & servParam(0)).ToString.Trim = "" Then
                        servParam(0) = _ObjRequest.UserCode
                    End If
                    servParam(1) = _ObjRequest.Password
                    servParam(2) = ("" & _ObjRequest.Category)
                    servParam(3) = ("" & _ObjRequest.Description)

                End If

            ElseIf Jenis = "XML" Then

                Dim doc As New XmlDocument

                doc.LoadXml(Parameter)
                Dim ObjXmlNodeList As XmlNodeList

                ObjXmlNodeList = doc.GetElementsByTagName("AppName")
                For Each ObjXmlNode As XmlNode In ObjXmlNodeList
                    AppName = ObjXmlNode.InnerText
                Next

                ObjXmlNodeList = doc.GetElementsByTagName("AppVersion")
                For Each ObjXmlNode As XmlNode In ObjXmlNodeList
                    AppVersion = ObjXmlNode.InnerText
                Next

                ObjXmlNodeList = doc.GetElementsByTagName("User")
                For Each ObjXmlNode As XmlNode In ObjXmlNodeList
                    servParam(0) = ObjXmlNode.InnerText
                Next
                If ("" & servParam(0)).ToString.Trim = "" Then
                    ObjXmlNodeList = doc.GetElementsByTagName("UserCode")
                    For Each ObjXmlNode As XmlNode In ObjXmlNodeList
                        servParam(0) = ObjXmlNode.InnerText
                    Next
                End If

                ObjXmlNodeList = doc.GetElementsByTagName("Password")
                For Each ObjXmlNode As XmlNode In ObjXmlNodeList
                    servParam(1) = ObjXmlNode.InnerText
                Next

            End If

            If Not servParam Is Nothing Then
                Dim tBalikan As Object() = serv.GetPackageList(AppName, AppVersion, servParam)
                balikan = ConvertArrayObjectToArrayString(tBalikan)

            Else
                'parameter tidak valid
                ReDim balikan(2)

                balikan(0) = "c10"
                balikan(1) = "Parameter tidak valid"
                balikan(2) = ""

            End If

        Catch ex As Exception

            ReDim balikan(2)
            balikan(0) = "c99"
            balikan(1) = "ERROR:" & ex.Message & ex.StackTrace
            balikan(2) = ""


        End Try

        '=====================================
        'PROSES CONVERT SESUAI DENGAN JENISNYA
        '=====================================

        sBalikan = ""

        If Jenis = "XML" Then

            sBalikan = "<result>" &
                        "<resultcode>" & balikan(0) & "</resultcode>" &
                        "<message>" & balikan(1) & "</message>" &
                        "<description>" & balikan(2) & "</description>" &
                        "</result>"

            If balikan(0) = "0" Then

                sBalikan = ConvertResultToXML(balikan, True)

            End If

        ElseIf Jenis = "JSON" Then

            sBalikan = "{" &
                        """resultcode"":""" & balikan(0) & """," &
                        """message"":""" & balikan(1) & """," &
                        """description"":""" & balikan(2) & """" &
                        "}"

            If balikan(0) = "0" Then

                sBalikan = ConvertResultToJSON(balikan, True)

            End If

        End If

        Return sBalikan

    End Function

    Public Function ConvertArrayObjectToArrayString(ByVal Parameter As Object()) As String()

        On Error Resume Next

        Dim Result As String()
        ReDim Result(Parameter.Length - 1)

        For i As Integer = 0 To Result.Length - 1
            Result(i) = ("" & Parameter(i)).ToString
        Next

        Return Result

    End Function

    Public Function ConvertResultToJSON(ByVal Parameter As Object(), Optional ByVal IsJSON As Boolean = False, Optional ByVal DataType() As String = Nothing) As String

        Dim Result As String = ""

        Dim dtParameter As DataTable = Nothing
        Dim JsonString As String = ""

        If IsJSON = False Then
            dtParameter = ConvertStringToDatatable(Parameter(2))
        Else
            JsonString = Parameter(2)
        End If

        Result = "{"
        Result &= """resultcode"":""" & Parameter(0) & ""","
        Result &= """message"":""" & Parameter(1) & ""","
        Result &= """description"":"

        If JsonString = "" Then
            Result &= "["
            For r As Integer = 0 To dtParameter.Rows.Count - 1
                Result &= "{"

                For c As Integer = 0 To dtParameter.Columns.Count - 1
                    Result &= """" & dtParameter.Columns(c).ColumnName.ToUpper & """:"
                    If DataType Is Nothing Then
                        Result &= """" & dtParameter.Rows(r).Item(c).ToString & """"
                    Else
                        If DataType(c) = "DOUBLE" Or DataType(c) = "INTEGER" Then
                            Result &= "" & dtParameter.Rows(r).Item(c).ToString & ""
                        Else
                            Result &= """" & dtParameter.Rows(r).Item(c).ToString & """"
                        End If
                    End If

                    If c <> dtParameter.Columns.Count - 1 Then
                        Result &= ","
                    End If
                Next

                Result &= "}"

                If r <> dtParameter.Rows.Count - 1 Then
                    Result &= ","
                End If

            Next
            Result &= "]"

        Else
            Result &= Parameter(2)

        End If

        Try
            If Not Parameter(3) Is Nothing Then

                dtParameter = Nothing
                JsonString = ""

                If IsJSON = False Then
                    dtParameter = ConvertStringToDatatable(Parameter(3))
                Else
                    JsonString = Parameter(3)
                End If

                Result &= ",""description2"":"

                If JsonString = "" Then
                    Result &= "["
                    For r As Integer = 0 To dtParameter.Rows.Count - 1
                        Result &= "{"

                        For c As Integer = 0 To dtParameter.Columns.Count - 1
                            Result &= """" & dtParameter.Columns(c).ColumnName.ToUpper & """:"
                            If DataType Is Nothing Then
                                Result &= """" & dtParameter.Rows(r).Item(c).ToString & """"
                            Else
                                If DataType(c) = "DOUBLE" Or DataType(c) = "INTEGER" Then
                                    Result &= "" & dtParameter.Rows(r).Item(c).ToString & ""
                                Else
                                    Result &= """" & dtParameter.Rows(r).Item(c).ToString & """"
                                End If
                            End If

                            If c <> dtParameter.Columns.Count - 1 Then
                                Result &= ","
                            End If
                        Next

                        Result &= "}"

                        If r <> dtParameter.Rows.Count - 1 Then
                            Result &= ","
                        End If

                    Next
                    Result &= "]"

                Else
                    Result &= Parameter(3)

                End If

            End If
        Catch ex As Exception
        End Try

        Result &= "}"

        Return Result

    End Function

    Public Function ConvertResultToXML(ByVal Parameter As Object(), Optional ByVal IsJSON As Boolean = False) As String

        Dim Result As String = ""

        Dim dtParameter As DataTable = Nothing
        Dim XmlString As String = ""

        If IsJSON = False Then
            dtParameter = ConvertStringToDatatable(Parameter(2))
        Else
            Dim JsonString As String = String.Format("{{ data: {0} }}", Parameter(2))

            Dim Header As String = "collection"
            Dim XmlDoc As XmlDocument = JsonConvert.DeserializeXmlNode(JsonString, Header)

            Dim sw As New IO.StringWriter()
            Dim xw As New XmlTextWriter(sw)
            XmlDoc.WriteTo(xw)
            XmlString = sw.ToString()

            XmlString = XmlString.Replace("<" & Header & ">", "").Replace("</" & Header & ">", "")
        End If

        Result = "<result>"
        Result &= "<resultcode>" & Parameter(0) & "</resultcode>"
        Result &= "<message>" & Parameter(1) & "</message>"
        Result &= "<description>"

        If XmlString = "" Then
            For r As Integer = 0 To dtParameter.Rows.Count - 1
                Result &= "<data>"

                For c As Integer = 0 To dtParameter.Columns.Count - 1
                    Result &= "<" & dtParameter.Columns(c).ColumnName.ToUpper & ">"
                    Result &= dtParameter.Rows(r).Item(c).ToString
                    Result &= "</" & dtParameter.Columns(c).ColumnName.ToUpper & ">"
                Next

                Result &= "</data>"
            Next

        Else
            Result &= XmlString

        End If

        Result &= "</description>"
        Result &= "</result>"

        Return Result

    End Function

    Private Function CertificateValidationCallBack(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
        Throw New NotImplementedException()
    End Function
End Class

