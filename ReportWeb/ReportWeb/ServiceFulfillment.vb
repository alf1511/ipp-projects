Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports MySql.Data.MySqlClient
Imports Newtonsoft.Json
Imports System.Collections.Generic
Imports Org.BouncyCastle.Crypto.Engines
Imports CrystalDecisions.[Shared].Json

<WebService(Namespace:="http://tempuri.org/")>
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Public Class ServiceFulfillment
    Inherits System.Web.Services.WebService

    Private MasterMCon As MySqlConnection
    Private SqlParam As New Dictionary(Of String, String)

    Private wsAPIKey As String
    Dim WService As New LocalCore
    Private mURLDO As String = ""
    Private mURLPO As String = ""
    Private mContentType As String = ""
    Private mTimeout As String = ""

    Private CustomHeaders As String() = Nothing

    Public Sub New()

        Dim ConSQL As String = "" & ConfigurationManager.AppSettings("ConSQL")

        If ConSQL <> "" Then
            'webconfig
            MasterMCon = New MySqlConnection(ConSQL)
        Else
            'default
            MasterMCon = New MySqlConnection("Server=localhost;Port=3306;Database=iexpress;Uid=root;Pwd=root;")
        End If

        mURLDO = ("" & ConfigurationManager.AppSettings("FulfillmentAPIDO")).Trim
        If mURLDO = "" Then
            mURLDO = "x"
        End If

        mURLPO = ("" & ConfigurationManager.AppSettings("FulfillmentAPIPO")).Trim
        If mURLPO = "" Then
            mURLPO = "x"
        End If

        wsAPIKey = ("" & ConfigurationManager.AppSettings("FulfillmentAPIKey")).Trim
        If wsAPIKey = "" Then
            wsAPIKey = "x"
        End If

        mContentType = "application/json"

        mTimeout = ("" & ConfigurationManager.AppSettings("FulfillmentAPITimeout")).Trim
        If mTimeout = "" Then
            mTimeout = "60"
        End If
        'ws.Timeout = mTimeout * 1000

    End Sub

    <WebMethod()>
    Public Function ChekConnection() As String
        Return "OK"
    End Function

    Public Class cResp
        Public resultcode As Integer = 1
        Public message As String
    End Class

    Public Class cReqDO
        Public do_number As String
        Public whouse_code As String
        Public supplier As String
        Public ecommerce As String
        Public ecommerce_no As String
        Public ecommerce_date As String
        Public expedition As String
        Public expedition_no As String
        Public expedition_date As String
        Public due_date As String
        Public expedition_service As String
        Public co_name As String
        Public co_address As String
        Public co_phone As String
        Public products As List(Of cReqProduct)
    End Class

    Public Class cReqProduct
        Public sku As String
        Public qty As String
    End Class

    <WebMethod()>
    Public Function FulfillmentDO(ByVal AppName As String, ByVal AppVersion As String, ByVal nLimit As Integer) As Object

        Dim Method As String = "FulfillmentDO"

        Dim Result As Boolean = False

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Try
            MCon = MasterMCon.Clone
            MCon.Open()

            Dim e As New ClsError
            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "START")


            Dim LimitData As Integer = 0
            Try
                LimitData = nLimit
            Catch ex As Exception
                LimitData = 0
            End Try

            If LimitData < 1 Then
                LimitData = 10
            End If

            Query = " Select DONumber, WHouseCode, Supplier"
            Query &= " , Ecommerce, EcommerceNo, cast(ifnull(date_format(EcommerceDate, '%Y-%m-%d %H:%i:%s'), '') as char) as EcommerceDate"
            Query &= " , Expedition, ExpeditionNo, cast(ifnull(date_format(ExpeditionDate, '%Y-%m-%d %H:%i:%s'), '') as char) as ExpeditionDate, ExpeditionService"
            Query &= " , cast(ifnull(date_format(DueDate, '%Y-%m-%d %H:%i:%s'), '') as char) as DueDate"
            Query &= " , CoName, CoAddress, CoPhone"
            Query &= " From Fulfillment_DO"
            Query &= " Where Status = '0'"
            Query &= " Limit " & LimitData

            Dim ObjSQL As New ClsSQL
            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
            If Not dtQuery Is Nothing Then

                If dtQuery.Rows.Count > 0 Then

                    For Each rowDO As DataRow In dtQuery.Rows

                        Dim ObjFungsi As New ClsFungsi

                        Dim ObjReqFulfillmentDO As New cReqDO

                        Dim ObjReqFulfillmentDOProduct As New List(Of cReqProduct)

                        ObjReqFulfillmentDO.do_number = rowDO.Item("DONumber").ToString
                        ObjReqFulfillmentDO.whouse_code = rowDO.Item("WhouseCode").ToString
                        ObjReqFulfillmentDO.supplier = rowDO.Item("Supplier").ToString
                        ObjReqFulfillmentDO.ecommerce = rowDO.Item("Ecommerce").ToString
                        ObjReqFulfillmentDO.ecommerce_no = rowDO.Item("EcommerceNo").ToString
                        ObjReqFulfillmentDO.ecommerce_date = rowDO.Item("EcommerceDate").ToString
                        ObjReqFulfillmentDO.expedition = rowDO.Item("Expedition").ToString
                        ObjReqFulfillmentDO.expedition_no = rowDO.Item("ExpeditionNo").ToString
                        ObjReqFulfillmentDO.expedition_date = rowDO.Item("ExpeditionDate").ToString
                        ObjReqFulfillmentDO.expedition_service = rowDO.Item("ExpeditionService").ToString
                        ObjReqFulfillmentDO.due_date = rowDO.Item("DueDate").ToString
                        ObjReqFulfillmentDO.co_name = rowDO.Item("CoName").ToString
                        ObjReqFulfillmentDO.co_address = rowDO.Item("CoAddress").ToString
                        ObjReqFulfillmentDO.co_phone = rowDO.Item("CoPhone").ToString

                        Query = " Select SKU, Qty"
                        Query &= " From Fulfillment_DOProduct"
                        Query &= " Where DONumber = @DONumber"

                        SqlParam = New Dictionary(Of String, String)
                        SqlParam.Add("@DONumber", rowDO.Item("DONumber").ToString)

                        Dim dtQueryProduct As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)
                        If Not dtQueryProduct Is Nothing Then

                            For Each rowProduct As DataRow In dtQueryProduct.Rows
                                Dim Product As New cReqProduct
                                Product.sku = rowProduct.Item("SKU")
                                Product.qty = rowProduct.Item("Qty")

                                ObjReqFulfillmentDOProduct.Add(Product)
                            Next

                        End If

                        ObjReqFulfillmentDO.products = ObjReqFulfillmentDOProduct

                        Dim Parameter As String = JsonConvert.SerializeObject(ObjReqFulfillmentDO)

                        Dim VendorApiKey As String = ""
                        Try
                            VendorApiKey = GetCredential(rowDO.Item("DONumber").ToString, "Fulfillment_DOProduct", "DONumber")
                        Catch ex As Exception
                            VendorApiKey = ""
                        End Try
                        If VendorApiKey = "" Then
                            VendorApiKey = wsAPIKey
                        End If

                        ReDim CustomHeaders(0)
                        CustomHeaders(0) = "X-API-KEY|" & VendorApiKey

                        Dim ParameterHeaders As String = ""
                        For i As Integer = 0 To CustomHeaders.Length - 1
                            If ParameterHeaders <> "" Then
                                ParameterHeaders &= ","
                            End If
                            ParameterHeaders &= CustomHeaders(i)
                        Next

                        Dim PushDOUrl As String = mURLDO

                        e.APIRequestLog(MCon, wsAppName, wsAppVersion, wsAppName & "." & Method, "", "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLDO, rowDO.Item("DONumber").ToString)

                        Dim Response As String = ""
                        Response = ObjFungsi.SendHTTP("", "", "", PushDOUrl, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
                        Response = ("" & Response).Trim

                        Dim ObjResponse As New cResp
                        Try
                            ObjResponse = JsonConvert.DeserializeObject(Of cResp)(Response)
                            If ObjResponse.resultcode = 0 And ObjResponse.message.ToLower.Contains("berhasil") Then

                                Query = " Update Fulfillment_DO"
                                Query &= " Set Status = '1'"
                                Query &= " Where DONumber = @DONumber"

                                SqlParam = New Dictionary(Of String, String)
                                SqlParam.Add("@DONumber", rowDO.Item("DONumber").ToString)

                                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                            Else

                                Query = " Update Fulfillment_DO"
                                Query &= " Set Status = '9'"
                                Query &= " , LastError = @LastError"
                                Query &= " Where DONumber = @DONumber"

                                SqlParam = New Dictionary(Of String, String)
                                SqlParam.Add("@DONumber", rowDO.Item("DONumber").ToString)
                                SqlParam.Add("@LastError", ObjResponse.message)

                                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                            End If
                        Catch ex As Exception
                            e.ErrorLog(MCon, wsAppName, wsAppVersion, Method, wsAppName, ex, Query)
                        End Try

                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, wsAppName & "." & Method, "", Response, mURLDO, rowDO.Item("DONumber").ToString)

                    Next

                End If

            End If

            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "FINISH")

            Result = True

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, wsAppName, ex, Query)

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



    Public Class cReqPO
        Public po_number As String
        Public whouse_code As String
        Public supplier As String
        Public products As List(Of cReqProduct)
    End Class

    <WebMethod()>
    Public Function FulfillmentPO(ByVal AppName As String, ByVal AppVersion As String, ByVal nLimit As Integer) As Object

        Dim Method As String = "FulfillmentPO"

        Dim Result As Boolean = False

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Try
            MCon = MasterMCon.Clone
            MCon.Open()

            Dim e As New ClsError
            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "START")


            Dim LimitData As Integer = 0
            Try
                LimitData = nLimit
            Catch ex As Exception
                LimitData = 0
            End Try

            If LimitData < 1 Then
                LimitData = 10
            End If

            Query = " Select PONumber, WHouseCode, Supplier"
            Query &= " From Fulfillment_PO"
            Query &= " Where Status = '0'"
            Query &= " Limit " & LimitData

            Dim ObjSQL As New ClsSQL
            Dim dtQuery As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, Nothing)
            If Not dtQuery Is Nothing Then

                If dtQuery.Rows.Count > 0 Then

                    For Each rowPO As DataRow In dtQuery.Rows

                        Dim ObjFungsi As New ClsFungsi

                        Dim ObjReqFulfillmentPO As New cReqPO

                        Dim ObjReqFulfillmentPOProduct As New List(Of cReqProduct)

                        ObjReqFulfillmentPO.po_number = rowPO.Item("PONumber").ToString
                        ObjReqFulfillmentPO.whouse_code = rowPO.Item("WhouseCode").ToString
                        ObjReqFulfillmentPO.supplier = rowPO.Item("Supplier").ToString

                        Query = " Select SKU, Qty"
                        Query &= " From Fulfillment_POProduct"
                        Query &= " Where PONumber = @PONumber"

                        SqlParam = New Dictionary(Of String, String)
                        SqlParam.Add("@PONumber", rowPO.Item("PONumber").ToString)

                        Dim dtQueryProduct As DataTable = ObjSQL.ExecDatatableWithParam(MCon, Query, SqlParam)
                        If Not dtQueryProduct Is Nothing Then

                            For Each rowProduct As DataRow In dtQueryProduct.Rows
                                Dim Product As New cReqProduct
                                Product.sku = rowProduct.Item("SKU")
                                Product.qty = rowProduct.Item("Qty")

                                ObjReqFulfillmentPOProduct.Add(Product)
                            Next

                        End If

                        ObjReqFulfillmentPO.products = ObjReqFulfillmentPOProduct

                        Dim Parameter As String = JsonConvert.SerializeObject(ObjReqFulfillmentPO)

                        Dim VendorApiKey As String = ""
                        Try
                            VendorApiKey = GetCredential(rowPO.Item("PONumber").ToString, "Fulfillment_POProduct", "PONumber")
                        Catch ex As Exception
                            VendorApiKey = ""
                        End Try
                        If VendorApiKey = "" Then
                            VendorApiKey = wsAPIKey
                        End If

                        ReDim CustomHeaders(0)
                        CustomHeaders(0) = "X-API-KEY|" & VendorApiKey

                        Dim ParameterHeaders As String = ""
                        For i As Integer = 0 To CustomHeaders.Length - 1
                            If ParameterHeaders <> "" Then
                                ParameterHeaders &= ","
                            End If
                            ParameterHeaders &= CustomHeaders(i)
                        Next

                        Dim PushPOUrl As String = mURLPO

                        e.APIRequestLog(MCon, wsAppName, wsAppVersion, wsAppName & "." & Method, "", "Header:" & ParameterHeaders & "-" & "Body:" & Parameter, mURLPO, rowPO.Item("PONumber").ToString)

                        Dim Response As String = ""
                        Response = ObjFungsi.SendHTTP("", "", "", PushPOUrl, Parameter, "", Encoding.Default, "60", "", mContentType, True, CustomHeaders)
                        Response = ("" & Response).Trim

                        Dim ObjResponse As New cResp
                        Try
                            ObjResponse = JsonConvert.DeserializeObject(Of cResp)(Response)
                            If ObjResponse.resultcode = 0 And ObjResponse.message.ToLower.Contains("berhasil") Then

                                Query = " Update Fulfillment_PO"
                                Query &= " Set Status = '1'"
                                Query &= " Where PONumber = @PONumber"

                                SqlParam = New Dictionary(Of String, String)
                                SqlParam.Add("@PONumber", rowPO.Item("PONumber").ToString)

                                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                            Else

                                Query = " Update Fulfillment_PO"
                                Query &= " Set Status = '9'"
                                Query &= " , LastError = @LastError"
                                Query &= " Where PONumber = @PONumber"

                                SqlParam = New Dictionary(Of String, String)
                                SqlParam.Add("@PONumber", rowPO.Item("PONumber").ToString)
                                SqlParam.Add("@LastError", ObjResponse.message)

                                ObjSQL.ExecNonQueryWithParam(MCon, Query, SqlParam)

                            End If
                        Catch ex As Exception
                            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, Query)
                        End Try

                        e.APIResponseLog(MCon, wsAppName, wsAppVersion, wsAppName & "." & Method, "", Response, mURLPO, rowPO.Item("PONumber").ToString)

                    Next

                End If

            End If

            e.DebugLog(MCon, AppName, AppVersion, Method, wsAppName, "FINISH")

            Result = True

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(AppName, AppVersion, Method, wsAppName, ex, Query)

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


    Public Function GetCredential(ByVal Number As String, ByVal TableName As String, ByVal ColumnName As String) As String

        Dim Method As String = "GetCredential"

        Dim Result As String = ""

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Try
            MCon = MasterMCon.Clone

            Dim e As New ClsError
            'e.DebugLog(wsAppName, wsAppVersion, Method, wsAppName, "START")

            Query = " Select a.api_key "
            Query &= " From " & TableName & " x "
            Query &= " Join "
            Query &= " ( "
            Query &= "   Select User "
            Query &= "   , SUBSTRING_INDEX(SUBSTRING_INDEX(addinfo, 'FULFILLVENDOR=', -1), ';', 1) AS vendor"
            Query &= "   From MstLogin"
            Query &= " ) l on x.AddUser = l.User "
            Query &= " Join Fulfillment_AccountCredential a on l.vendor = a.account "
            Query &= " Where x." & ColumnName & " = @Number "

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Number", Number)

            Result = ("" & ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)).ToString

            'e.DebugLog(wsAppName, wsAppVersion, Method, wsAppName, "FINISH")

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, Query)

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    Public Function GetPackageList(ByVal Number As String, ByVal TableName As String, ByVal ColumnName As String) As String

        Dim Method As String = "GetCredential"

        Dim Result As String = ""

        Dim MCon As New MySqlConnection
        Dim Query As String = ""

        Try
            MCon = MasterMCon.Clone

            Dim e As New ClsError
            'e.DebugLog(wsAppName, wsAppVersion, Method, wsAppName, "START")

            Query = " Select a.api_key "
            Query &= " From " & TableName & " x "
            Query &= " Join "
            Query &= " ( "
            Query &= "   Select User "
            Query &= "   , SUBSTRING_INDEX(SUBSTRING_INDEX(addinfo, 'FULFILLVENDOR=', -1), ';', 1) AS vendor"
            Query &= "   From MstLogin"
            Query &= " ) l on x.AddUser = l.User "
            Query &= " Join Fulfillment_AccountCredential a on l.vendor = a.account "
            Query &= " Where x." & ColumnName & " = @Number "

            Dim ObjSQL As New ClsSQL

            SqlParam = New Dictionary(Of String, String)
            SqlParam.Add("@Number", Number)

            Result = ("" & ObjSQL.ExecScalarWithParam(MCon, Query, SqlParam)).ToString

            'e.DebugLog(wsAppName, wsAppVersion, Method, wsAppName, "FINISH")

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(wsAppName, wsAppVersion, Method, wsAppName, ex, Query)

        Finally
            If MCon.State <> ConnectionState.Closed Then
                MCon.Close()
            End If
            Try
                MCon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function


    Public Function InputFulfillAccount(ByVal AppName As String, ByVal AppVersion As String _
    , ByVal Param As Object(), ByRef dsData As DataSet) As Object()

        Dim Method As String = "InputFulfillAccount"

        Dim Result As Object() = CreateResult()

        Dim Mcon As New MySqlConnection
        Dim Query As String = ""

        Dim User As String = Param(0).ToString
        Dim Password As String = Param(1).ToString

        Dim service As New LocalCore

        Try
            Dim dtData As DataTable = dsData.Tables(0).Copy

            Dim Mtrn As MySqlTransaction
            Mcon = MasterMCon.Clone

            Dim UserOK As Object() = service.ValidasiUser(Mcon, User, Password)
            If UserOK(0) <> "0" Then
                Result(1) = UserOK(1)
                GoTo Skip
            End If

            Dim Mcom As New MySqlCommand("", Mcon)
            Mcon.Open()

            If IsNothing(dtData) Then
                Result(1) = "Data Kosong"
                GoTo Skip
            End If

            Dim ObjSQL As New ClsSQL
            Dim ToolsInsertList = ""

            Query = ""
            Query = "INSERT IGNORE INTO Fulfillment_Account"
            Query &= "(account,name,alias,type,active_date,add_time,add_user,upd_time,upd_user)"
            Query &= "VALUES"

            SqlParam = New Dictionary(Of String, String)

            For i As Integer = 0 To dtData.Rows.Count - 1

                SqlParam.Add("@Account" & i, "" & dtData.Rows(i).Item("Account"))
                SqlParam.Add("@Name" & i, "" & dtData.Rows(i).Item("Name"))
                SqlParam.Add("@Alias" & i, "" & dtData.Rows(i).Item("Alias"))
                SqlParam.Add("@Type" & i, "" & dtData.Rows(i).Item("Type"))
                SqlParam.Add("@Active_Date" & i, "" & dtData.Rows(i).Item("Active_Date"))
                SqlParam.Add("@Add_Time" & i, "" & dtData.Rows(i).Item("Add_Time"))
                SqlParam.Add("@Add_User" & i, "" & dtData.Rows(i).Item("Add_User"))
                SqlParam.Add("@Upd_Time" & i, "" & dtData.Rows(i).Item("Upd_Time"))
                SqlParam.Add("@Upd_User" & i, "" & dtData.Rows(i).Item("Upd_User"))

                If i > 0 Then
                    Query &= ","
                End If
                Query &= "(@Account" & i & ",@Name" & i & ",@Alias" & i & ",@Type" & i & ","
                Query &= "@Active_Date" & i & ",@Add_Time" & i & ",@Add_User" & i & ",@Upd_Time" & i & ",@Upd_User" & i & ")"

            Next

            Query &= "ON DUPLICATE KEY UPDATE name = VALUES(Name), alias = VALUES(Alias)"
            Query &= ", type = VALUES(Type), active_Date = VALUES(Active_Date),"
            Query &= "upd_time = VALUES(Upd_Time), upd_user = VALUES(Upd_User)"

            If ObjSQL.ExecNonQueryWithParam(Mcon, Query, SqlParam) Then
                Result(0) = 0
                Result(1) = "Berhasil"
                Result(2) = ""
            Else
                Result(1) = "Gagal Query"
            End If
Skip:
        Catch ex As Exception
            Result(1) = ex.Message

        Finally

            If Not Mcon Is Nothing Then
                If Mcon.State <> ConnectionState.Closed Then
                    Mcon.Close()
                End If
                Mcon.Dispose()
            End If

        End Try

        Return Result
    End Function

    Dim wsAppName As String = "WServiceFulfillment"
    Dim wsAppVersion As String = "24.09.12.00"

    <WebMethod()>
    Public Function FulfillmentAddItem(ByVal AppName As String, ByVal AppVersion As String _
       , ByVal Param As Object(), ByRef dsData As DataSet) As Object()

        Dim Method As String = "FulfillmentAddItem"

        Dim Result As Object() = CreateResult()

        Dim Mcon As New MySqlConnection

        Dim Query As String = ""

        Dim LogKeyword As String = Format(Date.Now, "yyMMddHHmmss")

        'Dim User As String = Param(0).ToString
        Try
            'Dim Password As String = Param(1).ToString

            Mcon = MasterMCon.Clone
            Mcon.Open()

            'Dim UserOK As Object() = WService.ValidasiUser(Mcon, User, Password)
            'If UserOK(0) <> "0" Then
            '    Result(1) = UserOK(1)
            '    GoTo Skip
            'End If

            'WService.CreateRequestLog(AppName, AppVersion, Method, User, "", Param, LogKeyword)

            Dim dtData As DataTable = Nothing
            Try
                dtData = dsData.Tables(0).Copy
            Catch ex As Exception
                dtData = Nothing
            End Try
            If IsNothing(dtData) Then
                Result(1) = "Data Kosong"
                GoTo Skip
            End If

            SqlParam = New Dictionary(Of String, String)

            Query = ""
            Query &= " INSERT IGNORE INTO Fulfillment_Item ("
            Query &= " sku, vendor_code, supplier_code, name, add_time, add_user, active_date, inactive_date"
            Query &= " ) VALUES"

            For i As Integer = 0 To dtData.Rows.Count - 1

                If i > 0 Then
                    Query &= ","
                End If

                Query &= " ( @Sku" & i & ", @Vendor_Code" & i & ", @Supplier_Code" & i & ", @Name" & i
                Query &= " , now(), @Add_User, @Active_Date" & i & ", @Inactive_Date" & i
                Query &= " )"

                SqlParam.Add("@Sku" & i, "" & dtData.Rows(i).Item("Sku"))
                SqlParam.Add("@Vendor_Code" & i, "" & dtData.Rows(i).Item("Vendor_Code"))
                SqlParam.Add("@Supplier_Code" & i, "" & dtData.Rows(i).Item("Supplier_Code"))

                'If dtData.Rows(i).Item("Active_Date") = Nothing Then
                '    SqlParam.Add("@Active_Date" & i, Format(CDate("" & dtData.Rows(i).Item("Active_Date")), "yyyy-MM-dd"))
                'Else
                '    SqlParam.Add("@Active_Date" & i, Format(CDate("" & dtData.Rows(i).Item("Active_Date")), "yyyy-MM-dd"))
                'End If

                SqlParam.Add("@Name" & i, "" & dtData.Rows(i).Item("Name"))

                SqlParam.Add("@Active_Date" & i, Format(CDate("" & dtData.Rows(i).Item("Active_Date")), "yyyy-MM-dd"))
                If dtData.Rows(i).Item("Inactive_Date") = "" Then
                    SqlParam.Add("@Inactive_Date" & i, Nothing)
                Else
                    SqlParam.Add("@Inactive_Date" & i, "" & dtData.Rows(i).Item("Inactive_date"))
                End If

            Next

            Query &= " ON DUPLICATE KEY UPDATE"
            Query &= "   name = VALUES(Name), Supplier_Code = VALUES(Supplier_Code)"
            Query &= " , inactive_date = VALUES(Inactive_Date) ,upd_time = now(), upd_user = @Add_User"

            SqlParam.Add("@Add_User", Left(AppName & " " & AppVersion, 45))

            Dim ObjSQL As New ClsSQL

            If ObjSQL.ExecNonQueryWithParam(Mcon, Query, SqlParam) Then
                Result(0) = 0
                Result(1) = "OK"
                Result(2) = ""
            Else
                Result(1) = "Gagal Insert"
            End If
Skip:

            'WService.CreateResponseLog(AppName, AppVersion, Method, User, Result, , LogKeyword)

        Catch ex As Exception
            Dim e As New ClsError
            e.ErrorLog(Mcon, AppName, AppVersion, Method, User, ex, "", LogKeyword)

            Result(1) = ex.Message

        Finally
            Try
                Mcon.Close()
            Catch ex As Exception
            End Try
            Try
                Mcon.Dispose()
            Catch ex As Exception
            End Try

        End Try

        Return Result

    End Function

    <WebMethod()>
    Public Function A_WSDevelopment() As String
        Return "This is WS Development"
    End Function

    Private Function CreateResult() As Object()

        Dim Result As Object()
        ReDim Result(2)

        Result(0) = "9" 'Kode Proses
        Result(1) = "" 'Pesan hasil proses / pesan error
        Result(2) = Nothing 'Variable yang dikembalikan

        Return Result

    End Function

End Class
